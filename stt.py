import os
import struct
from io import BytesIO

import google.cloud.speech_v1p1beta1 as speech
import openai
import pvporcupine
import pyaudio
import win32com.client
from google.cloud import translate_v2

from slide_generator import generate_slide


def detect_question(text):
    question_words = [
        "who", "what", "when", "where", "why", "how",
        "are", "is", "am", "was", "were", "will", "would", "can", "could", "should", "shall",
        "have", "has", "had", "do", "does", "did", "might", "may", "must",
        "are there", "is there", "does it", "which", "whose", "whom",
        "can you", "could you", "will you", "would you", "should you", "shall you",
        "do you", "does he", "does she", "does one", "did you", "did he", "did she", "did one",
        "might you", "may you", "must you", "haven't", "hasn't", "hadn't", "don't", "doesn't", "didn't",
        "can't", "couldn't", "shouldn't", "shan't", "won't", "wouldn't", "mightn't", "mustn't", "aren't", "isn't",
        "weren't", "wasn't"
    ]

    return text.endswith("?") or any(text.lower().startswith(qw) for qw in question_words)


def transcribe_speech():
    client = speech.SpeechClient()
    config = speech.RecognitionConfig(
        encoding=speech.RecognitionConfig.AudioEncoding.LINEAR16,
        sample_rate_hertz=16000,
        language_code="he-IL",
    )

    audio_buffer = BytesIO()
    pa = pyaudio.PyAudio()
    stream = pa.open(
        rate=16000,
        channels=1,
        format=pyaudio.paInt16,
        input=True,
        frames_per_buffer=1024,
    )

    print("Recording...")

    try:
        for _ in range(0, int(16000 / 1024 * 3)):
            data = stream.read(1024, exception_on_overflow=False)
            audio_buffer.write(data)
    except IOError as e:
        print(f"Error while recording audio: {e}")
    finally:
        stream.stop_stream()
        stream.close()
        pa.terminate()

    audio_content = audio_buffer.getvalue()

    try:
        response = client.recognize(config=config, audio=speech.RecognitionAudio(content=audio_content))
        transcript = response.results[0].alternatives[0].transcript
        return transcript.strip()
    except Exception as e:
        print(f"Error during transcription: {e}")
        return ""


if __name__ == "__main__":
    with open("api_key.txt", "r") as f:
        openai.api_key = f.read().strip()

    # Start an instance of PowerPoint
    PowerPointApp = win32com.client.Dispatch("PowerPoint.Application")

    # Make the PowerPoint application visible
    PowerPointApp.Visible = True

    # Create a new presentation
    presentation = PowerPointApp.Presentations.Add()

    num_slides = 1

    os.environ[
        "GOOGLE_APPLICATION_CREDENTIALS"] = r"C:\Users\kfir1\AppData\Roaming\gcloud\application_default_credentials.json"

    porcupine = None
    pa = None
    audio_stream = None
    translate_client = translate_v2.Client()

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

        while True:
            pcm = audio_stream.read(porcupine.frame_length)
            pcm = struct.unpack_from("h" * porcupine.frame_length, pcm)
            result = porcupine.process(pcm)

            if result >= 0:
                print("Wakeword detected!")
                transcribed_text = transcribe_speech()

                translated_text = translate_client.translate(transcribed_text, target_language='en')['translatedText']

                if detect_question(translated_text):
                    print(f"Question: {translated_text}")
                    generate_slide(presentation, translated_text, num_slides)
                    num_slides += 1
                else:
                    print(f"Not a question {translated_text}")

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
