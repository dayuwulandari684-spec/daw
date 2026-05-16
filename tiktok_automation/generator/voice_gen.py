"""
Generator voiceover menggunakan Google TTS (gTTS).
Gratis, tidak perlu API key.
"""

import os
from pathlib import Path

from gtts import gTTS
from pydub import AudioSegment

import config


class VoiceGenerator:
    def __init__(self,
                 output_dir: str = config.AUDIO_DIR,
                 lang: str       = config.TTS_LANGUAGE,
                 tld: str        = config.TTS_TLD):
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.lang = lang
        self.tld  = tld

    def generate(self, text: str, filename: str = "voiceover") -> str:
        """
        Generate file MP3 dari teks.
        Return: path absolut ke file MP3.
        """
        mp3_path = self.output_dir / f"{filename}.mp3"

        print(f"[VOICE] Generating TTS → {mp3_path.name}")
        tts = gTTS(text=text, lang=self.lang, tld=self.tld, slow=False)
        tts.save(str(mp3_path))

        # Normalize volume (pydub)
        audio = AudioSegment.from_mp3(str(mp3_path))
        normalized = self._normalize(audio, target_dBFS=-14.0)
        normalized.export(str(mp3_path), format="mp3")

        duration = len(normalized) / 1000.0
        print(f"[VOICE] Selesai: {duration:.1f} detik")
        return str(mp3_path)

    def generate_sections(self, hook: str, problem: str,
                          solution: str, cta: str,
                          product_id: str) -> dict:
        """
        Generate voiceover per bagian.
        Return: dict {'hook': path, 'problem': path, ...}
        """
        sections = {
            "hook"    : hook,
            "problem" : problem,
            "solution": solution,
            "cta"     : cta,
        }
        paths = {}
        for name, text in sections.items():
            paths[name] = self.generate(text, f"{product_id}_{name}")
        return paths

    @staticmethod
    def _normalize(audio: AudioSegment, target_dBFS: float) -> AudioSegment:
        change = target_dBFS - audio.dBFS
        return audio.apply_gain(change)

    @staticmethod
    def get_duration_ms(mp3_path: str) -> int:
        """Durasi file audio dalam milidetik."""
        audio = AudioSegment.from_mp3(mp3_path)
        return len(audio)
