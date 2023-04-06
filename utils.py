from io import BytesIO

import google.cloud.speech_v1p1beta1 as speech
import pyaudio


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
