import os
import numpy as np
import re
import sherpa_onnx
import sounddevice as sd
import soundfile as sf
from scipy import signal

import sys
import struct
import threading
import pythoncom
from PySide6.QtWidgets import (
    QApplication, QWidget, QTextEdit, QVBoxLayout, QPushButton, QLabel, QComboBox, QHBoxLayout, QLineEdit, QSizePolicy,
    QSpinBox, QCheckBox, QDialog
)
from PySide6.QtCore import Signal, QObject
from PySide6.QtGui import QTextCursor, QColor, QIcon
import time

import shutil
import wave
import subprocess


class PlainTextEdit(QTextEdit):
    def insertFromMimeData(self, source):
        self.insertPlainText(source.text())


class WorkerSignals(QObject):
    lineChanged = Signal(int, str)
    finished = Signal()


class TTSWorker(threading.Thread):
    def __init__(self, app):
        super().__init__()

        self.app = app

        self.single_language_only = False

        self.repetition_times1 = 1
        self.repetition_times2 = 1

        self.last_target_edit = None
        self.last_position = 0
        self.last_length = 0
        self.last_background = None

        self._pause_event = threading.Event()
        self._stop_event = threading.Event()

    def pause(self):
        self._pause_event.clear()

    def resume(self):
        self._pause_event.set()

    def stop(self):
        self._stop_event.set()
        self._pause_event.set()

    def run(self):
        pythoncom.CoInitialize()

        total_lines = len(self.app.text1)

        while self.app.current_index < total_lines and not self._stop_event.is_set():
            while not self._pause_event.is_set() and not self._stop_event.is_set():
                pass

            if not self.single_language_only:
                self.app.signals.lineChanged.emit(self.app.current_index, "left")

            for i in range(self.repetition_times1):
                if not self.app.chinese_voice_selected1:
                    self.app.gen_config1.extra = {"lang": self.app.voice_code1}
                    self.app.gen_config1.sid = self.app.speaker_id1
                    self.app.gen_config1.speed = self.app.speed_factor

                    audio = self.app.tts1.generate(self.app.text1[self.app.current_index], self.app.gen_config1)
                else:
                    audio = self.app.tts2.generate(text=self.app.text1[self.app.current_index],
                                                   sid=self.app.speaker_id1, speed=self.app.speed_factor)

                sd.play(audio.samples, samplerate=audio.sample_rate)
                sd.wait()

                counter = 0
                start_time = time.time()
                while (time.time() - start_time) < self.app.pause1:
                    counter += 1

            if not self.single_language_only:
                self.app.signals.lineChanged.emit(self.app.current_index, "right")

                for i in range(self.repetition_times2):
                    self.app.gen_config1.extra = {"lang": self.app.voice_code2}
                    self.app.gen_config1.sid = self.app.speaker_id2
                    self.app.gen_config1.speed = self.app.speed_factor

                    if not self.app.chinese_voice_selected2:
                        audio = self.app.tts1.generate(self.app.text2[self.app.current_index], self.app.gen_config1)
                    else:
                        audio = self.app.tts2.generate(text=self.app.text2[self.app.current_index],
                                                       sid=self.app.speaker_id2, speed=self.app.speed_factor)

                    sd.play(audio.samples, samplerate=audio.sample_rate)
                    sd.wait()

                    counter = 0
                    start_time = time.time()
                    while (time.time() - start_time) < self.app.pause2:
                        counter += 1

            self.app.current_index += 1

        self.app.signals.finished.emit()


class TTSApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Language Meaning TTS Advanced")
        self.resize(960, 640)

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        SUPERTONIC_DIR = os.path.join(base_path, "tts_models", "sherpa-onnx-supertonic-3-tts-int8-2026-05-11")
        ZH_DIR = os.path.join(base_path, "tts_models", "vits-zh-aishell3")

        # Supertonic 3
        self.config1 = sherpa_onnx.OfflineTtsConfig(
            model=sherpa_onnx.OfflineTtsModelConfig(
                supertonic=sherpa_onnx.OfflineTtsSupertonicModelConfig(
                    duration_predictor=os.path.join(SUPERTONIC_DIR, "duration_predictor.int8.onnx"),
                    text_encoder=os.path.join(SUPERTONIC_DIR, "text_encoder.int8.onnx"),
                    vector_estimator=os.path.join(SUPERTONIC_DIR, "vector_estimator.int8.onnx"),
                    vocoder=os.path.join(SUPERTONIC_DIR, "vocoder.int8.onnx"),
                    tts_json=os.path.join(SUPERTONIC_DIR, "tts.json"),
                    unicode_indexer=os.path.join(SUPERTONIC_DIR, "unicode_indexer.bin"),
                    voice_style=os.path.join(SUPERTONIC_DIR, "voice.bin"),
                ),
                num_threads=2,
                provider="cpu",
            )
        )

        self.tts1 = sherpa_onnx.OfflineTts(self.config1)

        self.gen_config1 = sherpa_onnx.GenerationConfig()
        self.gen_config1.num_steps = 5

        # vits-zh-aishell3
        config_zh = sherpa_onnx.OfflineTtsConfig(
            model=sherpa_onnx.OfflineTtsModelConfig(
                vits=sherpa_onnx.OfflineTtsVitsModelConfig(
                    model=os.path.join(ZH_DIR, "vits-aishell3.onnx"),
                    lexicon=os.path.join(ZH_DIR, "lexicon.txt"),
                    tokens=os.path.join(ZH_DIR, "tokens.txt"),
                ),
                num_threads=2,
                provider="cpu",
            ),
            rule_fsts=",".join([
                os.path.join(ZH_DIR, f)
                for f in ["phone.fst", "date.fst", "number.fst", "new_heteronym.fst"]
            ]),
        )

        self.tts2 = sherpa_onnx.OfflineTts(config_zh)

        self.speed_factor = 1.0

        self.voice_name_list = [
            "Arabic (العربية)",
            "Bulgarian (Български)",
            "Chinese （中文）",
            "Croatian (Hrvatski)",
            "Czech (Čeština)",
            "Danish (Dansk)",
            "Dutch (Nederlands)",
            "English",
            "Estonian (Eesti)",
            "Finnish (Suomi)",
            "French (Français)",
            "German (Deutsch)",
            "Greek (Ελληνικά)",
            "Hindi (हिन्दी)",
            "Hungarian (Magyar)",
            "Indonesian (Bahasa Indonesia)",
            "Italian (Italiano)",
            "Japanese (日本語)",
            "Korean (한국어)",
            "Latvian (Latviešu)",
            "Lithuanian (Lietuvių)",
            "Polish (Polski)",
            "Portuguese (Português)",
            "Romanian (Română)",
            "Russian (Русский)",
            "Slovak (Slovenčina)",
            "Slovenian (Slovenščina)",
            "Spanish (Español)",
            "Swedish (Svenska)",
            "Turkish (Türkçe)",
            "Ukrainian (Українська)",
            "Vietnamese (Tiếng Việt)"
        ]

        self.voice_code_list = ["ar", "bg", "zh", "hr", "cs", "da", "nl", "en", "et", "fi", "fr", "de", "el", "hi",
                                "hu", "id", "it", "ja", "ko", "lv", "lt", "pl", "pt", "ro", "ru", "sk", "sl", "es",
                                "sv", "tr", "uk", "vi"]

        self.speaker_id_list = [str(i) for i in range(10)]

        self.chinese_speaker_id_list = [str(i) for i in range(174)]

        self.voice_code1 = ""
        self.voice_code2 = ""

        self.chinese_voice_selected1 = False
        self.chinese_voice_selected2 = False

        self.speaker_id1 = 0
        self.speaker_id2 = 0

        self.text1 = ""
        self.text2 = ""

        self.pause1 = 0.0
        self.pause2 = 0.0

        self.layout = QVBoxLayout(self)

        first_row_layout = QHBoxLayout(self)
        first_row_layout.addStretch()
        self.about_btn = QPushButton("ℹ️ About")
        self.about_btn.setFlat(True)
        self.about_btn.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.about_btn.clicked.connect(self.open_license_window)
        first_row_layout.addWidget(self.about_btn)

        self.layout.addLayout(first_row_layout)

        description_label = QLabel(
            "<b>When outputting audio files, the system will generate:</b><br>"
            "• <code>FirstVoiceTemporary12345999</code> folder<br>"
            "• <code>SecondVoiceTemporary12345999</code> folder<br>"
            "• <code>SilentTemporary12345999.wav</code><br>"
            "<font color='#2980B9'><b>Note:</b> Avoid renaming any of your personal files to the names above "
            "to prevent your personal data from being deleted.</font>"
        )
        self.layout.addWidget(description_label)

        self.single_language_checkbox = QCheckBox(
            "Single language only (if enabled, works like a regular single-language TTS, otherwise, alternates between two languages)")
        self.single_language_checkbox.stateChanged.connect(self.on_single_language_checkbox_changed)
        self.layout.addWidget(self.single_language_checkbox)

        language_global_layout = QHBoxLayout()
        language_global_layout.setContentsMargins(0, 4, 0, 0)

        # Language 1 voice selection
        language_vertical_layout1 = QVBoxLayout()
        language_vertical_layout1.setContentsMargins(0, 0, 0, 10)

        voice_layout = QHBoxLayout()
        voice_label = QLabel("First Voice:")
        voice_layout.addWidget(voice_label)

        self.voice_combo1 = QComboBox()
        self.voice_combo1.addItems(self.voice_name_list)
        self.voice_combo1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.voice_combo1.currentIndexChanged.connect(self.on_voice1_change)
        voice_layout.addWidget(self.voice_combo1)

        voice_label = QLabel("Speaker Id:")
        voice_layout.addWidget(voice_label)

        self.speaker_id_combo1 = QComboBox()
        self.speaker_id_combo1.addItems(self.speaker_id_list)
        self.speaker_id_combo1.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.speaker_id_combo1.setFixedWidth(80)
        self.speaker_id_combo1.currentIndexChanged.connect(self.on_speaker_id1_change)
        voice_layout.addWidget(self.speaker_id_combo1)

        self.font_combo1 = QComboBox()
        self.font_options = ["aaa", "Aa", "AA"]
        self.font_sizes = [9, 12, 16]
        self.font_combo1.addItems(self.font_options)
        self.font_combo1.setFixedWidth(80)
        self.font_combo1.currentIndexChanged.connect(self.change_font_size1)
        voice_layout.addWidget(self.font_combo1)

        language_vertical_layout1.addLayout(voice_layout)

        repetition_layout1 = QHBoxLayout()

        self.repetition_label1 = QLabel("Repeat: ")
        repetition_layout1.addWidget(self.repetition_label1)

        self.repetition_spinbox1 = QSpinBox()
        self.repetition_spinbox1.setRange(1, 100)
        self.repetition_spinbox1.setValue(1)
        self.repetition_spinbox1.setFixedWidth(100)
        repetition_layout1.addWidget(self.repetition_spinbox1)

        self.repetition_label2 = QLabel(" times")
        repetition_layout1.addWidget(self.repetition_label2)

        repetition_layout1.addStretch(1)
        language_vertical_layout1.addLayout(repetition_layout1)

        pause_layout1 = QHBoxLayout()

        repetition_label3 = QLabel("Pause: ")
        pause_layout1.addWidget(repetition_label3)

        self.pause_edit1 = QLineEdit("0.0")
        self.pause_edit1.textChanged.connect(self.update_pause1)
        pause_layout1.addWidget(self.pause_edit1)

        self.repetition_label4 = QLabel(" seconds - each line")
        pause_layout1.addWidget(self.repetition_label4)

        pause_layout1.addStretch(1)
        language_vertical_layout1.addLayout(pause_layout1)

        # Text edit 1
        self.text_edit1 = PlainTextEdit()
        language_vertical_layout1.addWidget(self.text_edit1)
        self.original_background1 = self.text_edit1.palette().base().color()
        self.text_edit1.textChanged.connect(self.text_changed)

        language_global_layout.addLayout(language_vertical_layout1)

        # Language 2 voice selection
        self.language_vertical_layout_widget2 = QWidget()
        language_vertical_layout2 = QVBoxLayout(self.language_vertical_layout_widget2)
        language_vertical_layout2.setContentsMargins(10, 0, 0, 10)

        voice_layout = QHBoxLayout()
        voice_label = QLabel("Second Voice:")
        voice_layout.addWidget(voice_label)

        self.voice_combo2 = QComboBox()

        self.voice_combo2.addItems(self.voice_name_list)

        self.voice_combo2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.voice_combo2.currentIndexChanged.connect(self.on_voice2_change)
        voice_layout.addWidget(self.voice_combo2)

        voice_label = QLabel("Speaker Id:")
        voice_layout.addWidget(voice_label)

        self.speaker_id_combo2 = QComboBox()
        self.speaker_id_combo2.addItems(self.speaker_id_list)
        self.speaker_id_combo2.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.speaker_id_combo2.setFixedWidth(80)
        self.speaker_id_combo2.currentIndexChanged.connect(self.on_speaker_id2_change)
        voice_layout.addWidget(self.speaker_id_combo2)

        self.font_combo2 = QComboBox()
        self.font_options = ["aaa", "Aa", "AA"]
        self.font_sizes = [9, 12, 16]
        self.font_combo2.addItems(self.font_options)
        self.font_combo2.setFixedWidth(80)
        self.font_combo2.currentIndexChanged.connect(self.change_font_size2)
        voice_layout.addWidget(self.font_combo2)

        language_vertical_layout2.addLayout(voice_layout)

        repetition_layout2 = QHBoxLayout()

        repetition_label5 = QLabel("Repeat: ")
        repetition_layout2.addWidget(repetition_label5)

        self.repetition_spinbox2 = QSpinBox()
        self.repetition_spinbox2.setRange(1, 100)
        self.repetition_spinbox2.setValue(1)
        self.repetition_spinbox2.setFixedWidth(100)
        repetition_layout2.addWidget(self.repetition_spinbox2)

        repetition_label6 = QLabel(" times")
        repetition_layout2.addWidget(repetition_label6)

        repetition_layout2.addStretch(1)
        language_vertical_layout2.addLayout(repetition_layout2)

        pause_layout2 = QHBoxLayout()

        repetition_label7 = QLabel("Pause: ")
        pause_layout2.addWidget(repetition_label7)

        self.pause_edit2 = QLineEdit("0.0")
        self.pause_edit2.textChanged.connect(self.update_pause2)
        pause_layout2.addWidget(self.pause_edit2)

        repetition_label8 = QLabel(" seconds - each line")
        pause_layout2.addWidget(repetition_label8)

        pause_layout2.addStretch(1)
        language_vertical_layout2.addLayout(pause_layout2)

        # Text edit 2
        self.text_edit2 = PlainTextEdit()
        language_vertical_layout2.addWidget(self.text_edit2)
        self.original_background2 = self.text_edit2.palette().base().color()
        self.text_edit2.textChanged.connect(self.text_changed)

        language_global_layout.addWidget(self.language_vertical_layout_widget2)

        self.layout.addLayout(language_global_layout)

        # Controls
        control_layout = QHBoxLayout()

        self.play_btn = QPushButton("▶️ Play")
        self.play_btn.clicked.connect(self.toggle_play)
        control_layout.addWidget(self.play_btn)

        control_layout.addWidget(QLabel("Speed:"))
        self.speed_edit = QLineEdit("1.0")
        control_layout.addWidget(self.speed_edit)
        self.speed_edit.textChanged.connect(self.update_speed)

        self.save_btn = QPushButton("💾 Save")
        self.save_btn.clicked.connect(self.save_file)
        control_layout.addWidget(self.save_btn)

        self.layout.addLayout(control_layout)

        # TTS worker
        self.worker = None
        self.signals = WorkerSignals()
        self.signals.lineChanged.connect(self.highlight_line)
        self.signals.finished.connect(self.on_finished)

        self.is_playing = False

        self.lines = []
        self.text_modified = True
        self.current_index = 0

    def on_single_language_checkbox_changed(self, state):
        if state == 2:
            self.language_vertical_layout_widget2.hide()

            self.repetition_label1.hide()
            self.repetition_spinbox1.hide()
            self.repetition_label2.hide()
            self.repetition_label4.setText(" seconds - each sentence")
        else:
            self.language_vertical_layout_widget2.show()

            self.repetition_label1.show()
            self.repetition_spinbox1.show()
            self.repetition_label2.show()
            self.repetition_label4.setText(" seconds - each line")

    def on_voice1_change(self, index):
        self.voice_code1 = self.voice_code_list[index]

        if not self.chinese_voice_selected1:
            if index == 2:
                self.chinese_voice_selected1 = True
                self.speaker_id_combo1.clear()
                self.speaker_id_combo1.addItems(self.chinese_speaker_id_list)
        else:
            if index != 2:
                self.chinese_voice_selected1 = False
                self.speaker_id_combo1.clear()
                self.speaker_id_combo1.addItems(self.speaker_id_list)

    def on_voice2_change(self, index):
        self.voice_code2 = self.voice_code_list[index]

        if not self.chinese_voice_selected2:
            if index == 2:
                self.chinese_voice_selected2 = True
                self.speaker_id_combo2.clear()
                self.speaker_id_combo2.addItems(self.chinese_speaker_id_list)
        else:
            if index != 2:
                self.chinese_voice_selected2 = False
                self.speaker_id_combo2.clear()
                self.speaker_id_combo2.addItems(self.speaker_id_list)

    def on_speaker_id1_change(self, index):
        self.speaker_id1 = index

    def on_speaker_id2_change(self, index):
        self.speaker_id2 = index

    def change_font_size1(self, index):
        font = self.text_edit1.font()
        font.setPointSize(self.font_sizes[index])
        self.text_edit1.setFont(font)

    def change_font_size2(self, index):
        font = self.text_edit2.font()
        font.setPointSize(self.font_sizes[index])
        self.text_edit2.setFont(font)

    def update_speed(self):
        try:
            self.speed_factor = float(self.speed_edit.text())
        except ValueError:
            self.speed_edit.setText("1.0")
            self.speed_factor = 1.0

    def update_pause1(self):
        try:
            self.pause1 = round(float(self.pause_edit1.text()), 1)
        except ValueError:
            self.pause_edit1.setText("0.0")
            self.pause1 = 0.0

    def update_pause2(self):
        try:
            self.pause2 = round(float(self.pause_edit2.text()), 1)
        except ValueError:
            self.pause_edit2.setText("0.0")
            self.pause2 = 0.0

    def toggle_play(self):
        text_lines = self.text_edit1.toPlainText().splitlines()
        if self.text_modified or not self.lines:
            self.lines = text_lines
            self.text_modified = False
            self.current_index = 0

        if not self.lines:
            return

        if not self.is_playing:
            self.worker = TTSWorker(self)

            self.worker.single_language_only = self.single_language_checkbox.isChecked()

            self.worker.repetition_times1 = self.repetition_spinbox1.value()
            self.worker.repetition_times2 = self.repetition_spinbox2.value()

            self.worker._pause_event.set()
            self.worker.start()
            self.is_playing = True
            self.play_btn.setText("⏸️ Pause")
        else:
            if self.worker._pause_event.is_set():
                self.worker.pause()
                self.play_btn.setText("▶️ Play")

                if self.current_index > 0:
                    self.current_index -= 1
            else:
                self.worker.resume()
                self.play_btn.setText("⏸️ Pause")

    def on_finished(self):
        self.is_playing = False
        self.play_btn.setText("▶️ Play")
        self.current_index = 0

    def highlight_line(self, idx, side):
        last_target_edit = getattr(self, 'last_target_edit', None)

        if last_target_edit is not None:
            last_target_edit.blockSignals(True)
            prev_cursor = last_target_edit.textCursor()

            last_position = getattr(self, 'last_position', 0)
            last_length = getattr(self, 'last_length', 0)
            last_background = getattr(self, 'last_background', None)

            prev_cursor.setPosition(last_position)
            prev_cursor.setPosition(last_position + last_length, QTextCursor.KeepAnchor)
            last_target_edit.setTextCursor(prev_cursor)
            last_target_edit.setTextBackgroundColor(last_background)

            prev_cursor.clearSelection()
            last_target_edit.setTextCursor(prev_cursor)
            last_target_edit.blockSignals(False)

        if side == "left":
            target_edit = self.text_edit1
            lines_pool = self.text1
            current_bg = self.original_background1
        else:
            target_edit = self.text_edit2
            lines_pool = self.text2
            current_bg = self.original_background2

        if idx >= len(lines_pool):
            return

        pos = 0
        for i in range(idx):
            pos += len(lines_pool[i]) + 1
        line_len = len(lines_pool[idx])

        target_edit.blockSignals(True)
        cursor = target_edit.textCursor()

        cursor.setPosition(pos)
        cursor.setPosition(pos + line_len, QTextCursor.KeepAnchor)
        target_edit.setTextCursor(cursor)
        target_edit.setTextBackgroundColor(QColor("green"))

        target_edit.ensureCursorVisible()

        cursor.clearSelection()
        target_edit.setTextCursor(cursor)
        target_edit.blockSignals(False)

        self.last_target_edit = target_edit
        self.last_position = pos
        self.last_length = line_len
        self.last_background = current_bg

    def save_file(self):
        self.save_btn.setEnabled(False)
        self.save_btn.setText("⏳ Saving")

        self.generate_01s_silent_wav()
        silent_01s_path = "SilentTemporary12345999.wav"
        silent_01s_samples, _ = sf.read(silent_01s_path, dtype='float32')

        QApplication.processEvents()

        current_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()

        # Create folders
        dir1 = os.path.join(current_dir, "FirstVoiceTemporary12345999")
        dir2 = os.path.join(current_dir, "SecondVoiceTemporary12345999")
        for folder in [dir1, dir2]:
            if os.path.exists(folder):
                shutil.rmtree(folder)
            os.makedirs(folder, exist_ok=True)

        max_len = max(len(self.text1), len(self.text2))
        self.text1 += [""] * (max_len - len(self.text1))

        if not self.single_language_checkbox.isChecked():
            self.text2 += [""] * (max_len - len(self.text2))

        voice_code1 = self.voice_code_list[self.voice_combo1.currentIndex()]
        voice_code2 = self.voice_code_list[self.voice_combo2.currentIndex()]

        all_audio_segments = []
        wav_files = []

        # Calculate the time of pause 1
        pause_repetition_count1 = int(self.pause1 / 0.1)

        if self.single_language_checkbox.isChecked():
            for idx, f_line in enumerate(self.text1, 1):
                if f_line.strip():
                    f_wav = os.path.join(dir1, f"F{idx}.wav")

                    if not self.chinese_voice_selected1:
                        self.gen_config1.extra = {"lang": self.voice_code1}
                        self.gen_config1.sid = self.speaker_id1
                        self.gen_config1.speed = 1.0

                        audio = self.tts1.generate(f_line, self.gen_config1)
                        audio_samples = np.array(audio.samples, dtype=np.float32)
                    else:
                        audio = self.tts2.generate(text=f_line, sid=self.speaker_id1, speed=1.0)

                        audio_samples = np.array(audio.samples, dtype=np.float32)

                        # Resample from 8000 to 44100
                        audio_samples = signal.resample(audio.samples, int(len(audio_samples) * 5.5125)).astype(
                            np.float32)

                    sf.write(f_wav, audio_samples, self.tts1.sample_rate)
                    wav_files.append(f_wav)

                    all_audio_segments.append(audio_samples)

                    for _ in range(pause_repetition_count1):
                        all_audio_segments.append(silent_01s_samples)
        else:
            # Calculate the time of pause 2
            pause_repetition_count2 = int(self.pause2 / 0.1)

            for idx, (f_line, c_line) in enumerate(zip(self.text1, self.text2), 1):
                if f_line.strip():
                    f_wav = os.path.join(dir1, f"F{idx}.wav")

                    if not self.chinese_voice_selected1:
                        self.gen_config1.extra = {"lang": self.voice_code1}
                        self.gen_config1.sid = self.speaker_id1
                        self.gen_config1.speed = 1.0

                        audio = self.tts1.generate(f_line, self.gen_config1)
                        audio_samples = np.array(audio.samples, dtype=np.float32)
                    else:
                        audio = self.tts2.generate(text=f_line, sid=self.speaker_id1, speed=1.0)

                        audio_samples = np.array(audio.samples, dtype=np.float32)

                        # Resample from 8000 to 44100
                        audio_samples = signal.resample(audio.samples, int(len(audio_samples) * 5.5125)).astype(
                            np.float32)

                    sf.write(f_wav, audio_samples, self.tts1.sample_rate)
                    wav_files.append(f_wav)

                    for _ in range(self.repetition_spinbox1.value()):
                        all_audio_segments.append(audio_samples)

                        for _ in range(pause_repetition_count1):
                            all_audio_segments.append(silent_01s_samples)

                if c_line.strip():
                    c_wav = os.path.join(dir2, f"C{idx}.wav")

                    if not self.chinese_voice_selected2:
                        self.gen_config1.extra = {"lang": self.voice_code2}
                        self.gen_config1.sid = self.speaker_id2
                        self.gen_config1.speed = 1.0

                        audio = self.tts1.generate(c_line, self.gen_config1)
                        audio_samples = np.array(audio.samples, dtype=np.float32)
                    else:
                        audio = self.tts2.generate(text=c_line, sid=self.speaker_id2, speed=1.0)

                        audio_samples = np.array(audio.samples, dtype=np.float32)

                        # Resample from 8000 to 44100
                        audio_samples = signal.resample(audio.samples, int(len(audio_samples) * 5.5125)).astype(
                            np.float32)

                    sf.write(c_wav, audio_samples, self.tts1.sample_rate)
                    wav_files.append(c_wav)

                    for _ in range(self.repetition_spinbox2.value()):
                        all_audio_segments.append(audio_samples)

                        for _ in range(pause_repetition_count2):
                            all_audio_segments.append(silent_01s_samples)

        if not wav_files:
            return

        # Combine wav
        combined_path = os.path.join(current_dir, "Combined.wav")
        combined_audio = np.concatenate(all_audio_segments)
        sf.write(combined_path, combined_audio, self.tts1.sample_rate)

        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = current_dir

        if sys.platform == "win32":
            ffmpeg_path = os.path.join(base_path, "ffmpeg-8.0-essentials_build", "bin", "ffmpeg.exe")
        else:
            ffmpeg_path = os.path.join(base_path, "ffmpeg-8.0-essentials_build", "bin", "ffmpeg")

        output_ogg = os.path.join(current_dir, "Combined.ogg")

        subprocess.run([
            ffmpeg_path,
            "-y",
            "-i", combined_path,
            "-filter:a", f"atempo={self.speed_factor}",
            output_ogg
        ])

        if os.path.exists(combined_path):
            os.remove(combined_path)
        for folder in [dir1, dir2]:
            if os.path.exists(folder):
                shutil.rmtree(folder)

        if os.path.exists(silent_01s_path):
            os.remove(silent_01s_path)

        self.save_btn.setEnabled(True)
        self.save_btn.setText("💾 Save")

    def open_license_window(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("About & Licenses")
        dialog.resize(700, 520)

        layout = QVBoxLayout(dialog)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(
            r"""The icon was generated by ChatGPT.

This software uses the following open-source components and models:

1. sherpa-onnx
   - Source URL: https://github.com/k2-fsa/sherpa-onnx
   - License: Apache License 2.0

2. Supertonic 3 (TTS Model)
   - Source URL: https://huggingface.co/Supertone/supertonic-3
   - License: BigScience Open RAIL-M License (https://huggingface.co/Supertone/supertonic-3/blob/main/LICENSE)
   - Notice: Audio generated by this model must not be used for impersonation, fraud, or unlawful purposes. Publicly distributed audio must be disclosed as AI-generated.

3. vits-zh-aishell3 (TTS Model)
   - Source URL: https://huggingface.co/csukuangfj/vits-zh-aishell3/tree/main
   - License: Apache License 2.0 (https://huggingface.co/datasets/choosealicense/licenses/blob/main/markdown/apache-2.0.md)

======================================================================

BigScience Open RAIL-M License
dated August 18, 2022

Section I: PREAMBLE

This Open RAIL-M License was created by BigScience, a collaborative open innovation project aimed at
the responsible development and use of large multilingual datasets and Large Language Models
(“LLMs”). While a similar license was originally designed for the BLOOM model, we decided to adapt it
and create this license in order to propose a general open and responsible license applicable to other
machine learning based AI models (e.g. multimodal generative models).
In short, this license strives for both the open and responsible downstream use of the accompanying
model. When it comes to the open character, we took inspiration from open source permissive licenses
regarding the grant of IP rights. Referring to the downstream responsible use, we added use-based
restrictions not permitting the use of the Model in very specific scenarios, in order for the licensor to be
able to enforce the license in case potential misuses of the Model may occur. Even though downstream
derivative versions of the model could be released under different licensing terms, the latter will always
have to include - at minimum - the same use-based restrictions as the ones in the original license (this
license).
The development and use of artificial intelligence (“AI”), does not come without concerns. The world has
witnessed how AI techniques may, in some instances, become risky for the public in general. These risks
come in many forms, from racial discrimination to the misuse of sensitive information.
BigScience believes in the intersection between open and responsible AI development; thus, this License
aims to strike a balance between both in order to enable responsible open-science in the field of AI.
This License governs the use of the model (and its derivatives) and is informed by the model card
associated with the model.

NOW THEREFORE, You and Licensor agree as follows:

1. Definitions
(a) "License" means the terms and conditions for use, reproduction, and Distribution as defined in
this document.
(b) “Data” means a collection of information and/or content extracted from the dataset used with the
Model, including to train, pretrain, or otherwise evaluate the Model. The Data is not licensed under
this License.
(c)“Output” means the results of operating a Model as embodied in informational content resulting
therefrom.
(d)“Model” means any accompanying machine-learning based assemblies (including checkpoints),
consisting of learnt weights, parameters (including optimizer states), corresponding to the model
architecture as embodied in the Complementary Material, that have been trained or tuned, in whole or
in part on the Data, using the Complementary Material.
(e) “Derivatives of the Model” means all modifications to the Model, works based on the Model, or any
other model which is created or initialized by transfer of patterns of the weights, parameters,
activations or output of the Model, to the other model, in order to cause the other model to perform
similarly to the Model, including - but not limited to - distillation methods entailing the use of
intermediate data representations or methods based on the generation of synthetic data by the Model
for training the other model.
(f)“Complementary Material” means the accompanying source code and scripts used to define,
run, load, benchmark or evaluate the Model, and used to prepare data for training or evaluation, if
any. This includes any accompanying documentation, tutorials, examples, etc, if any.
(g) “Distribution” means any transmission, reproduction, publication or other sharing of the Model or
Derivatives of the Model to a third party, including providing the Model as a hosted service made
available by electronic or other remote means - e.g. API-based or web access.
(h) “Licensor” means the copyright owner or entity authorized by the copyright owner that is
granting the License, including the persons or entities that may have rights in the Model and/or
distributing the Model.
(i) "You" (or "Your") means an individual or Legal Entity exercising permissions granted by this
License and/or making use of the Model for whichever purpose and in any field of use, including
usage of the Model in an end-use application - e.g. chatbot, translator, image generator.
(j) “Third Parties” means individuals or legal entities that are not under common control with
Licensor or You.
(k) "Contribution" means any work of authorship, including the original version of the Model and
any modifications or additions to that Model or Derivatives of the Model thereof, that is
intentionally submitted to Licensor for inclusion in the Model by the copyright owner or by an
individual or Legal Entity authorized to submit on behalf of the copyright owner. For the
purposes of this definition,
“submitted” means any form of electronic, verbal, or written
communication sent to the Licensor or its representatives, including but not limited to
communication on electronic mailing lists, source code control systems, and issue tracking
systems that are managed by, or on behalf of, the Licensor for the purpose of discussing and
improving the Model, but excluding communication that is conspicuously marked or otherwise
designated in writing by the copyright owner as "Not a Contribution."
(l) "Contributor" means Licensor and any individual or Legal Entity on behalf of whom a
Contribution has been received by Licensor and subsequently incorporated within the Model.


Section II: INTELLECTUAL PROPERTY RIGHTS

Both copyright and patent grants apply to the Model, Derivatives of the Model and Complementary
Material. The Model and Derivatives of the Model are subject to additional terms as described in Section III.

2. Grant of Copyright License. Subject to the terms and conditions of this License, each Contributor
hereby grants to You a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable copyright license to reproduce, prepare, publicly display, publicly perform, sublicense, and distribute the
Complementary Material, the Model, and Derivatives of the Model.

3. Grant of Patent License. Subject to the terms and conditions of this License and where and as
applicable, each Contributor hereby grants to You a perpetual, worldwide, non-exclusive, no-charge,
royalty-free, irrevocable (except as stated in this paragraph) patent license to make, have made, use, offer
to sell, sell, import, and otherwise transfer the Model and the Complementary Material, where such
license applies only to those patent claims licensable by such Contributor that are necessarily infringed by
their Contribution(s) alone or by combination of their Contribution(s) with the Model to which such
Contribution(s) was submitted. If You institute patent litigation against any entity (including a cross-claim
or counterclaim in a lawsuit) alleging that the Model and/or Complementary Material or a Contribution
incorporated within the Model and/or Complementary Material constitutes direct or contributory patent
infringement, then any patent licenses granted to You under this License for the Model and/or Work shall
terminate as of the date such litigation is asserted or filed.
Section III: CONDITIONS OF USAGE, DISTRIBUTION AND REDISTRIBUTION

4. Distribution and Redistribution. You may host for Third Party remote access purposes (e.g.
software-as-a-service), reproduce and distribute copies of the Model or Derivatives of the Model thereof
in any medium, with or without modifications, provided that You meet the following conditions:

a. Use-based restrictions as referenced in paragraph 5 MUST be included as an enforceable provision
by You in any type of legal agreement (e.g. a license) governing the use and/or distribution of the
Model or Derivatives of the Model, and You shall give notice to subsequent users You Distribute to,
that the Model or Derivatives of the Model are subject to paragraph 5. This provision does not apply
to the use of Complementary Material.

b. You must give any Third Party recipients of the Model or Derivatives of the Model a copy of this
License;

c. You must cause any modified files to carry prominent notices stating that You changed the files;

d. You must retain all copyright, patent, trademark, and attribution notices excluding those notices
that do not pertain to any part of the Model, Derivatives of the Model.
You may add Your own copyright statement to Your modifications and may provide additional or
different license terms and conditions - respecting paragraph 4.a.
- for use, reproduction, or Distribution
of Your modifications, or for any such Derivatives of the Model as a whole, provided Your use,
reproduction, and Distribution of the Model otherwise complies with the conditions stated in this License.

5. Use-based restrictions. The restrictions set forth in Attachment A are considered Use-based restrictions.
Therefore You cannot use the Model and the Derivatives of the Model for the specified restricted uses. You
may use the Model subject to this License, including only for lawful purposes and in accordance with the
License. Use may include creating any content with, finetuning, updating, running, training, evaluating and/or
reparametrizing the Model. You shall require all of Your users who use the Model or a Derivative of the Model
to comply with the terms of this paragraph (paragraph 5).

6. The Output You Generate. Except as set forth herein, Licensor claims no rights in the Output You
generate using the Model. You are accountable for the Output you generate and its subsequent uses. No
use of the output can contravene any provision as stated in the License.

Section IV: OTHER PROVISIONS

7. Updates and Runtime Restrictions. To the maximum extent permitted by law, Licensor reserves the
right to restrict (remotely or otherwise) usage of the Model in violation of this License, update the Model
through electronic means, or modify the Output of the Model based on updates. You shall undertake
reasonable efforts to use the latest version of the Model.

8. Trademarks and related. Nothing in this License permits You to make use of Licensors’ trademarks,
trade names, logos or to otherwise suggest endorsement or misrepresent the relationship between the
parties; and any rights not expressly granted herein are reserved by the Licensors.

9. Disclaimer of Warranty. Unless required by applicable law or agreed to in writing, Licensor provides
the Model and the Complementary Material (and each Contributor provides its Contributions) on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied,
including, without limitation, any warranties or conditions of TITLE, NON-INFRINGEMENT,
MERCHANTABILITY , or FITNESS FOR A PARTICULAR PURPOSE. You are solely responsible for
determining the appropriateness of using or redistributing the Model, Derivatives of the Model, and the
Complementary Material and assume any risks associated with Your exercise of permissions under this
License.

10. Limitation of Liability. In no event and under no legal theory, whether in tort (including negligence),
contract, or otherwise, unless required by applicable law (such as deliberate and grossly negligent acts) or
agreed to in writing, shall any Contributor be liable to You for damages, including any direct, indirect,
special, incidental, or consequential damages of any character arising as a result of this License or out of
the use or inability to use the Model and the Complementary Material (including but not limited to
damages for loss of goodwill, work stoppage, computer failure or malfunction, or any and all other
commercial damages or losses), even if such Contributor has been advised of the possibility of such
damages.

11. Accepting Warranty or Additional Liability. While redistributing the Model, Derivatives of the
Model and the Complementary Material thereof, You may choose to offer, and charge a fee for, acceptance
of support, warranty, indemnity, or other liability obligations and/or rights consistent with this License.
However, in accepting such obligations, You may act only on Your own behalf and on Your sole
responsibility, not on behalf of any other Contributor, and only if You agree to indemnify, defend, and
hold each Contributor harmless for any liability incurred by, or claims asserted against, such Contributor
by reason of your accepting any such warranty or additional liability.

12. If any provision of this License is held to be invalid, illegal or unenforceable, the remaining
provisions shall be unaffected thereby and remain valid as if such provision had not been set forth herein.

END OF TERMS AND CONDITIONS

Attachment A

Use Restrictions

You agree not to use the Model or Derivatives of the Model:
(a) In any way that violates any applicable national, federal, state, local or international law
or regulation;
(b) For the purpose of exploiting, harming or attempting to exploit or harm minors in any
way;
(c) To generate or disseminate verifiably false information and/or content with the purpose of
harming others;
(d) To generate or disseminate personal identifiable information that can be used to harm an
individual;
(e) To generate or disseminate information and/or content (e.g. images, code, posts, articles),
and place the information and/or content in any context (e.g. bot generating tweets)
without expressly and intelligibly disclaiming that the information and/or content is
machine generated;
(f) To defame, disparage or otherwise harass others;
(g) To impersonate or attempt to impersonate (e.g. deepfakes) others without their consent;
(h) For fully automated decision making that adversely impacts an individual’s legal rights or
otherwise creates or modifies a binding, enforceable obligation;
(i) For any use intended to or which has the effect of discriminating against or harming
individuals or groups based on online or offline social behavior or known or predicted
personal or personality characteristics;
(j) To exploit any of the vulnerabilities of a specific group of persons based on their age,
social, physical or mental characteristics, in order to materially distort the behavior of a
person pertaining to that group in a manner that causes or is likely to cause that person or
another person physical or psychological harm;
(k) For any use intended to or which has the effect of discriminating against individuals or
groups based on legally protected characteristics or categories;
(l) To provide medical advice and medical results interpretation;
(m) To generate or disseminate information for the purpose to be used for administration of
justice, law enforcement, immigration or asylum processes, such as predicting an
individual will commit fraud/crime commitment (e.g. by text profiling, drawing causal
relationships between assertions made in documents, indiscriminate and
arbitrarily-targeted use).

======================================================================

                             Apache License
                       Version 2.0, January 2004
                    http://www.apache.org/licenses/
TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION

1. Definitions.

"License" shall mean the terms and conditions for use, reproduction, and distribution as defined by Sections 1 through 9 of this document.

"Licensor" shall mean the copyright owner or entity authorized by the copyright owner that is granting the License.

"Legal Entity" shall mean the union of the acting entity and all other entities that control, are controlled by, or are under common control with that entity. For the purposes of this definition, "control" means (i) the power, direct or indirect, to cause the direction or management of such entity, whether by contract or otherwise, or (ii) ownership of fifty percent (50%) or more of the outstanding shares, or (iii) beneficial ownership of such entity.

"You" (or "Your") shall mean an individual or Legal Entity exercising permissions granted by this License.

"Source" form shall mean the preferred form for making modifications, including but not limited to software source code, documentation source, and configuration files.

"Object" form shall mean any form resulting from mechanical transformation or translation of a Source form, including but not limited to compiled object code, generated documentation, and conversions to other media types.

"Work" shall mean the work of authorship, whether in Source or Object form, made available under the License, as indicated by a copyright notice that is included in or attached to the work (an example is provided in the Appendix below).

"Derivative Works" shall mean any work, whether in Source or Object form, that is based on (or derived from) the Work and for which the editorial revisions, annotations, elaborations, or other modifications represent, as a whole, an original work of authorship. For the purposes of this License, Derivative Works shall not include works that remain separable from, or merely link (or bind by name) to the interfaces of, the Work and Derivative Works thereof.

"Contribution" shall mean any work of authorship, including the original version of the Work and any modifications or additions to that Work or Derivative Works thereof, that is intentionally submitted to Licensor for inclusion in the Work by the copyright owner or by an individual or Legal Entity authorized to submit on behalf of the copyright owner. For the purposes of this definition, "submitted" means any form of electronic, verbal, or written communication sent to the Licensor or its representatives, including but not limited to communication on electronic mailing lists, source code control systems, and issue tracking systems that are managed by, or on behalf of, the Licensor for the purpose of discussing and improving the Work, but excluding communication that is conspicuously marked or otherwise designated in writing by the copyright owner as "Not a Contribution."

"Contributor" shall mean Licensor and any individual or Legal Entity on behalf of whom a Contribution has been received by Licensor and subsequently incorporated within the Work.

2. Grant of Copyright License. Subject to the terms and conditions of this License, each Contributor hereby grants to You a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable copyright license to reproduce, prepare Derivative Works of, publicly display, publicly perform, sublicense, and distribute the Work and such Derivative Works in Source or Object form.

3. Grant of Patent License. Subject to the terms and conditions of this License, each Contributor hereby grants to You a perpetual, worldwide, non-exclusive, no-charge, royalty-free, irrevocable (except as stated in this section) patent license to make, have made, use, offer to sell, sell, import, and otherwise transfer the Work, where such license applies only to those patent claims licensable by such Contributor that are necessarily infringed by their Contribution(s) alone or by combination of their Contribution(s) with the Work to which such Contribution(s) was submitted. If You institute patent litigation against any entity (including a cross-claim or counterclaim in a lawsuit) alleging that the Work or a Contribution incorporated within the Work constitutes direct or contributory patent infringement, then any patent licenses granted to You under this License for that Work shall terminate as of the date such litigation is filed.

4. Redistribution. You may reproduce and distribute copies of the Work or Derivative Works thereof in any medium, with or without modifications, and in Source or Object form, provided that You meet the following conditions:

(a) You must give any other recipients of the Work or Derivative Works a copy of this License; and (b) You must cause any modified files to carry prominent notices stating that You changed the files; and (c) You must retain, in the Source form of any Derivative Works that You distribute, all copyright, patent, trademark, and attribution notices from the Source form of the Work, excluding those notices that do not pertain to any part of the Derivative Works; and (d) If the Work includes a "NOTICE" text file as part of its distribution, then any Derivative Works that You distribute must include a readable copy of the attribution notices contained within such NOTICE file, excluding those notices that do not pertain to any part of the Derivative Works, in at least one of the following places: within a NOTICE text file distributed as part of the Derivative Works; within the Source form or documentation, if provided along with the Derivative Works; or, within a display generated by the Derivative Works, if and wherever such third-party notices normally appear. The contents of the NOTICE file are for informational purposes only and do not modify the License. You may add Your own attribution notices within Derivative Works that You distribute, alongside or as an addendum to the NOTICE text from the Work, provided that such additional attribution notices cannot be construed as modifying the License. You may add Your own copyright statement to Your modifications and may provide additional or different license terms and conditions for use, reproduction, or distribution of Your modifications, or for any such Derivative Works as a whole, provided Your use, reproduction, and distribution of the Work otherwise complies with the conditions stated in this License.

5. Submission of Contributions. Unless You explicitly state otherwise, any Contribution intentionally submitted for inclusion in the Work by You to the Licensor shall be under the terms and conditions of this License, without any additional terms or conditions. Notwithstanding the above, nothing herein shall supersede or modify the terms of any separate license agreement you may have executed with Licensor regarding such Contributions.

6. Trademarks. This License does not grant permission to use the trade names, trademarks, service marks, or product names of the Licensor, except as required for reasonable and customary use in describing the origin of the Work and reproducing the content of the NOTICE file.

7. Disclaimer of Warranty. Unless required by applicable law or agreed to in writing, Licensor provides the Work (and each Contributor provides its Contributions) on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied, including, without limitation, any warranties or conditions of TITLE, NON-INFRINGEMENT, MERCHANTABILITY, or FITNESS FOR A PARTICULAR PURPOSE. You are solely responsible for determining the appropriateness of using or redistributing the Work and assume any risks associated with Your exercise of permissions under this License.

8. Limitation of Liability. In no event and under no legal theory, whether in tort (including negligence), contract, or otherwise, unless required by applicable law (such as deliberate and grossly negligent acts) or agreed to in writing, shall any Contributor be liable to You for damages, including any direct, indirect, special, incidental, or consequential damages of any character arising as a result of this License or out of the use or inability to use the Work (including but not limited to damages for loss of goodwill, work stoppage, computer failure or malfunction, or any and all other commercial damages or losses), even if such Contributor has been advised of the possibility of such damages.

9. Accepting Warranty or Additional Liability. While redistributing the Work or Derivative Works thereof, You may choose to offer, and charge a fee for, acceptance of support, warranty, indemnity, or other liability obligations and/or rights consistent with this License. However, in accepting such obligations, You may act only on Your own behalf and on Your sole responsibility, not on behalf of any other Contributor, and only if You agree to indemnify, defend, and hold each Contributor harmless for any liability incurred by, or claims asserted against, such Contributor by reason of your accepting any such warranty or additional liability.

END OF TERMS AND CONDITIONS

APPENDIX: How to apply the Apache License to your work.

  To apply the Apache License to your work, attach the following
  boilerplate notice, with the fields enclosed by brackets "[]"
  replaced with your own identifying information. (Don't include
  the brackets!)  The text should be enclosed in the appropriate
  comment syntax for the file format. We also recommend that a
  file or class name and description of purpose be included on the
  same "printed page" as the copyright notice for easier
  identification within third-party archives.
Copyright 2026 JusticeForSure2026

Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the License at

   http://www.apache.org/licenses/LICENSE-2.0
Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.
"""
        )

        layout.addWidget(text_edit)
        dialog.exec()

    def closeEvent(self, event):
        if self.worker:
            self.worker.stop()
        event.accept()

    def text_changed(self):
        self.text_modified = True

        if self.single_language_checkbox.isChecked():
            self.text1 = [s.strip() for s in re.split(r'[。.!?]', self.text_edit1.toPlainText()) if s.strip()]
        else:
            self.text1 = self.text_edit1.toPlainText().replace("/", ",").splitlines()
            self.text2 = self.text_edit2.toPlainText().replace("/", ",").splitlines()

        self.current_index = 0

    def generate_01s_silent_wav(self):
        output_path = "SilentTemporary12345999.wav"

        num_samples = int(self.tts1.sample_rate * 0.1)

        silent_data = struct.pack('<' + 'h' * num_samples, *(0 for _ in range(num_samples)))

        with wave.open(output_path, 'wb') as wav_file:
            wav_file.setnchannels(1)
            wav_file.setsampwidth(2)
            wav_file.setframerate(self.tts1.sample_rate)
            wav_file.writeframes(silent_data)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    if getattr(sys, 'frozen', False):
        base_icon_path = sys._MEIPASS
    else:
        base_icon_path = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()

    if sys.platform == "win32":
        icon_path = os.path.join(base_icon_path, "language_meaning_tts_advanced_icon.ico")
        app.setWindowIcon(QIcon(icon_path))

    elif sys.platform == "darwin":
        icon_path = os.path.join(base_icon_path, "language_meaning_tts_advanced_icon.png")
        app.setWindowIcon(QIcon(icon_path))

    window = TTSApp()
    window.show()
    sys.exit(app.exec())
