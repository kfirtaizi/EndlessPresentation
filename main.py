import html
import os
import struct
import uuid

import openai
import pvporcupine
import pyaudio
import win32com.client
from google.cloud import translate_v2

from slide_generator import generate_slide
import collections
import collections.abc
from pptx import Presentation
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


def generate_realtime_slide(presentation, translated_text):
    temp_presentation_path = os.path.abspath(f"temp_presentation-{str(uuid.uuid4())}.pptx")

    temp_presentation = Presentation()
    generate_slide(temp_presentation, translated_text)
    temp_presentation.save(temp_presentation_path)

    opened_temp_presentation = PowerPointApp.Presentations.Open(temp_presentation_path)
    opened_temp_presentation.Slides(1).Copy()
    presentation.Slides.Paste()
    opened_temp_presentation.Close()
    os.remove(temp_presentation_path)


if __name__ == "__main__":
    configure_api_keys()

    PowerPointApp, presentation = configure_presentation()

    porcupine = None
    pa = None
    audio_stream = None
    translate_client = translate_v2.Client()
    num_slides = 1

    try:
        wakeword_model_path = "Bambino_it_windows_v2_2_0.ppn"
        model_path = "porcupine_params_it.pv"
        porcupine = pvporcupine.create(keyword_paths=[wakeword_model_path],
                                       model_path=model_path,
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

                translated_text = html.unescape(
                    translate_client.translate(transcribed_text, target_language='en')['translatedText'])

                if detect_question(translated_text):
                    print(f"Question: {translated_text}")
                    generate_realtime_slide(presentation, translated_text)
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
