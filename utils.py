from collections import defaultdict
from io import BytesIO
from operator import itemgetter

import google.cloud.speech_v1p1beta1 as speech
import openai
import pyaudio
from PIL import Image


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
        "weren't", "wasn't", "what is", "what's", "what are", "what're", "what was", "what's", "what were", "what've",
        "what have", "what has", "what had", "what do", "what does", "what did", "what can", "what could",
        "what should",
        "what shall", "what will", "what would", "what might", "what may", "what must"
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


def rgb_to_int(color):
    r, g, b = color
    return r + (g * 256) + (b * 256 * 256)


def get_dominant_colors(image_path, num_colors=3):
    img = Image.open(image_path).resize((150, 150), Image.ANTIALIAS)
    pixels = img.getcolors(img.size[0] * img.size[1])

    color_count = defaultdict(int)
    for count, color in pixels:
        color_count[color] += count

    sorted_colors = sorted(color_count.items(), key=itemgetter(1), reverse=True)
    return [color for color, count in sorted_colors[:num_colors]]


def contrast_color(color):
    r, g, b = color
    return 255 - r, 255 - g, 255 - b


def ask_chatgpt(prompt, max_tokens=1024):
    response = openai.Completion.create(
        engine="text-davinci-003",
        prompt=prompt,
        temperature=0.7,
        max_tokens=max_tokens,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0,
    )

    return response.choices[0].text
