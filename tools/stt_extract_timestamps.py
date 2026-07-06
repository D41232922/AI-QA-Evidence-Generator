#!/usr/bin/env python
"""
Extract timestamps where the word 'screenshot' (case-insensitive) occurs using Vosk.
Outputs one timestamp per line (seconds, float).

Requires: pip install vosk soundfile
Also requires ffmpeg available on PATH.

Usage: stt_extract_timestamps.py <video-file>
"""
import sys
import json
import subprocess
import os
import wave

try:
    from vosk import Model, KaldiRecognizer
except Exception as e:
    sys.stderr.write("Missing vosk. Install with: pip install vosk\n")
    sys.exit(2)

if len(sys.argv) < 2:
    print("Usage: stt_extract_timestamps.py <video-file>")
    sys.exit(1)

video = sys.argv[1]
# temporary wav file
wav = os.path.join(os.getcwd(), "__stt_audio.wav")
# extract mono 16k wav
cmd = ["ffmpeg", "-y", "-i", video, "-ac", "1", "-ar", "16000", "-vn", "-f", "wav", wav]
ret = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
if ret.returncode != 0:
    sys.stderr.write("ffmpeg audio extraction failed:\n" + ret.stderr.decode())
    sys.exit(3)

# load model
model_path = os.environ.get('VOSK_MODEL', 'model')
if not os.path.exists(model_path):
    sys.stderr.write(f"Vosk model not found at '{model_path}'. Set VOSK_MODEL env or place model folder.\n")
    sys.exit(4)

wf = wave.open(wav, "rb")
if wf.getnchannels() != 1 or wf.getsampwidth() != 2 or wf.getframerate() != 16000:
    sys.stderr.write("Audio file must be PCM 16k mono\n")
    wf.close()
    os.remove(wav)
    sys.exit(5)

model = Model(model_path)
rec = KaldiRecognizer(model, wf.getframerate())
rec.SetWords(True)

timestamps = []
while True:
    data = wf.readframes(4000)
    if len(data) == 0:
        break
    if rec.AcceptWaveform(data):
        res = json.loads(rec.Result())
        for w in res.get('result', []):
            if 'screenshot' in w.get('word', '').lower():
                timestamps.append(w.get('start'))
# final
res = json.loads(rec.FinalResult())
for w in res.get('result', []):
    if 'screenshot' in w.get('word', '').lower():
        timestamps.append(w.get('start'))

wf.close()
try:
    os.remove(wav)
except Exception:
    pass

# print unique sorted
for t in sorted(set(timestamps)):
    print(f"{t:.3f}")
