import html
import os
import struct

import openai
import pvporcupine
import pyaudio
import win32com.client
from google.cloud import translate_v2

from slide_generator import generate_slide
from utils import transcribe_speech, detect_question


def configure_api_keys():
    with open("keys/api_key.txt", "r") as f:
        openai.api_key = f.read().strip()

    os.environ[
        "GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\kfir1\AppData\Roaming\gcloud\application_default_credentials.json"


def configure_presentation():
    # Start an instance of PowerPoint
    PowerPointApp = win32com.client.Dispatch("PowerPoint.Application")

    # Make the PowerPoint application visible
    PowerPointApp.Visible = True

    # Create a new presentation
    presentation = PowerPointApp.Presentations.Add()

    return PowerPointApp, presentation


if __name__ == "__main__":
    configure_api_keys()

    PowerPointApp, presentation = configure_presentation()

    porcupine = None
    pa = None
    audio_stream = None
    translate_client = translate_v2.Client()
    num_slides = 1

    try:
        porcupine = pvporcupine.create(keywords=["bumblebee"],
                                       access_key='aVBwJU8YoxqCExAuddswdVuNace1HGaEWRzN9e3T1hGVZjewetbaFA==')
        pa = pyaudio.PyAudio()
        audio_stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length,
        )

        print("Waiting for a wakeword...")
        while True:
            pcm = audio_stream.read(porcupine.frame_length)
            pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
            result = porcupine.process(pcm)

            if result >= 0:
                print("Wakeword detected!")
                transcribed_text = transcribe_speech()

                translated_text = html.unescape(translate_client.translate(transcribed_text, target_language='en')['translatedText'])

                if detect_question(translated_text):
                    print(f"Question: {translated_text}")
                    generate_slide(presentation, translated_text, num_slides)
                    num_slides += 1
                else:
                    print(f"Not a question {translated_text}")

                print("Waiting for a wakeword...")

    finally:
        if porcupine is not None:
            porcupine.delete()

        if audio_stream is not None:
            audio_stream.close()

        if pa is not None:
            pa.terminate()

        # Save the presentation
        presentation.SaveAs(os.path.join(os.getcwd(), "real_time_presentation.pptx"))

        # Close the presentation
        presentation.Close()

        # Quit the PowerPoint application
        PowerPointApp.Quit()
