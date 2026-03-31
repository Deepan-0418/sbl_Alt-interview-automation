import sys
import json
import pandas as pd
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QListWidget, QLabel, QTextEdit,
    QMessageBox, QProgressBar, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QFont

class ConverterThread(QThread):
    log_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, json_files, output_path):
        super().__init__()
        self.json_files = json_files
        self.output_path = output_path

    def run(self):
        try:
            self.log_signal.emit("🚀 Starting conversion of {} files...".format(len(self.json_files)))
            
            summary_data = []
            hourly_data = []
            breaks_data = []

            for i, json_file in enumerate(self.json_files):
                self.log_signal.emit(f"📄 Processing {i+1}/{len(self.json_files)}: {json_file.split('/')[-1]}")
                self.progress_signal.emit(int((i + 1) / len(self.json_files) * 100))

                with open(json_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # === SUMMARY ROW (one per session) ===
                summary_row = {
                    'username': data.get('username', ''),
                    'system_name': data.get('system_name', ''),
                    'session_start': data.get('session_start', ''),
                    'session_end': data.get('session_end', ''),
                    'snapshot_time': data.get('snapshot_time', ''),
                    'session_duration_seconds': data.get('session_duration_seconds', 0),
                    'session_duration_formatted': data.get('session_duration_formatted', ''),
                    'total_duration_seconds': data.get('total_duration_seconds', 0),
                    'total_duration_formatted': data.get('total_duration_formatted', ''),
                    'save_reason': data.get('save_reason', ''),
                    'save_timestamp': data.get('save_timestamp', ''),
                    'break_count': data.get('break_count', 0),
                }

                # Flatten keystroke_counts
                kc = data.get('keystroke_counts', {})
                summary_row.update({
                    'keystroke_total': kc.get('total', 0),
                    'keystroke_character': kc.get('character_keys', 0),
                    'keystroke_space': kc.get('space_keys', 0),
                    'keystroke_backspace': kc.get('backspace_keys', 0),
                    'keystroke_enter': kc.get('enter_keys', 0),
                    'keystroke_modifier': kc.get('modifier_keys', 0),
                    'keystroke_special': kc.get('special_keys', 0),
                })

                # Flatten kpm_metrics
                kpm = data.get('kpm_metrics', {})
                summary_row.update({
                    'kpm_current': kpm.get('current_kpm', 0),
                    'kpm_peak': kpm.get('peak_kpm', 0),
                    'kpm_average': kpm.get('average_kpm', 0),
                })

                # Flatten time_metrics
                tm = data.get('time_metrics', {})
                summary_row.update({
                    'time_active_seconds': tm.get('active_seconds', 0),
                    'time_idle_seconds': tm.get('idle_seconds', 0),
                    'time_active_formatted': tm.get('active_formatted', ''),
                    'time_idle_formatted': tm.get('idle_formatted', ''),
                    'typing_efficiency_percent': tm.get('typing_efficiency_percent', 0),
                })

                summary_data.append(summary_row)

                # === SESSION IDENTIFIER for linking sheets ===
                session_id = f"{data.get('username', 'unknown')}_{data.get('session_start', 'unknown')}"

                # === HOURLY BREAKDOWN ===
                for hour_entry in data.get('hourly_breakdown', []):
                    hourly_data.append({
                        'session_id': session_id,
                        'username': data.get('username', ''),
                        'system_name': data.get('system_name', ''),
                        'session_start': data.get('session_start', ''),
                        'hour': hour_entry.get('hour', ''),
                        'date': hour_entry.get('date', ''),
                        'total_keys': hour_entry.get('total_keys', 0),
                        'character_keys': hour_entry.get('character_keys', 0),
                        'space_keys': hour_entry.get('space_keys', 0),
                        'backspace_keys': hour_entry.get('backspace_keys', 0),
                        'enter_keys': hour_entry.get('enter_keys', 0),
                        'modifier_keys': hour_entry.get('modifier_keys', 0),
                        'special_keys': hour_entry.get('special_keys', 0),
                        'active_seconds': hour_entry.get('active_seconds', 0),
                    })

                # === BREAKS ===
                for break_entry in data.get('breaks', []):
                    breaks_data.append({
                        'session_id': session_id,
                        'username': data.get('username', ''),
                        'system_name': data.get('system_name', ''),
                        'session_start': data.get('session_start', ''),
                        'break_start': break_entry.get('start_time', ''),
                        'break_end': break_entry.get('end_time', ''),
                        'reason': break_entry.get('reason', ''),
                        'duration_seconds': break_entry.get('duration_seconds', 0),
                        'duration_formatted': break_entry.get('duration_formatted', ''),
                    })

            # Create DataFrames
            summary_df = pd.DataFrame(summary_data)
            hourly_df = pd.DataFrame(hourly_data)
            breaks_df = pd.DataFrame(breaks_data)

            # Nice column order
            summary_cols = [
                'username', 'system_name', 'session_start', 'session_end',
                'session_duration_formatted', 'total_duration_formatted',
                'keystroke_total', 'keystroke_character', 'keystroke_space',
                'keystroke_backspace', 'keystroke_enter', 'keystroke_modifier',
                'keystroke_special', 'kpm_current', 'kpm_peak', 'kpm_average',
                'time_active_seconds', 'time_idle_seconds', 'typing_efficiency_percent',
                'break_count', 'save_reason'
            ]
            summary_df = summary_df[summary_cols + [c for c in summary_df.columns if c not in summary_cols]]

            # Write Excel with formatting
            with pd.ExcelWriter(self.output_path, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                hourly_df.to_excel(writer, sheet_name='Hourly', index=False)
                breaks_df.to_excel(writer, sheet_name='Breaks', index=False)

                # Auto-adjust column widths
                for sheet_name in ['Summary', 'Hourly', 'Breaks']:
                    worksheet = writer.sheets[sheet_name]
                    for col in worksheet.columns:
                        max_length = 0
                        column = col[0].column_letter
                        for cell in col:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column].width = adjusted_width

            self.log_signal.emit("✅ Conversion completed successfully!")
            self.finished_signal.emit(True, f"Excel file saved:\n{self.output_path}")

        except Exception as e:
            self.log_signal.emit(f"❌ ERROR: {str(e)}")
            self.finished_signal.emit(False, str(e))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🔥 Keyboard Activity → Excel Converter")
        self.setMinimumSize(1000, 720)
        self.setStyleSheet("""
            QMainWindow { background: #0f0f1a; }
            QLabel { color: #fff; font-size: 15px; }
            QListWidget { background: #1a1a2e; border: 2px solid #00d4ff; border-radius: 12px; color: #fff; }
            QTextEdit { background: #1a1a2e; color: #00ffaa; font-family: SF Mono; border-radius: 12px; }
            QPushButton { 
                background: #00d4ff; 
                color: #000; 
                font-weight: bold; 
                padding: 14px 24px; 
                border-radius: 9999px; 
            }
            QPushButton:hover { background: #00ffdd; }
        """)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setSpacing(20)
        layout.setContentsMargins(30, 30, 30, 30)

        # Title
        title = QLabel("📊 Multi-User Keyboard Activity JSON Converter")
        title.setFont(QFont("Inter", 22, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # File selection
        self.select_btn = QPushButton("📁 Select Multiple JSON Files")
        self.select_btn.setIcon(QIcon.fromTheme("document-open"))
        self.select_btn.clicked.connect(self.select_files)
        layout.addWidget(self.select_btn)

        self.file_list = QListWidget()
        layout.addWidget(QLabel("Selected JSON files:"))
        layout.addWidget(self.file_list)

        # Output
        hbox = QHBoxLayout()
        self.output_btn = QPushButton("💾 Choose Output Excel File")
        self.output_btn.clicked.connect(self.select_output)
        hbox.addWidget(self.output_btn)
        
        self.output_label = QLabel("No output file selected yet")
        self.output_label.setStyleSheet("color: #aaa; padding: 12px; background: #1a1a2e; border-radius: 12px;")
        hbox.addWidget(self.output_label, stretch=1)
        layout.addLayout(hbox)

        # Process button
        self.process_btn = QPushButton("🚀 CONVERT TO EXCEL")
        self.process_btn.setStyleSheet("font-size: 18px; padding: 18px; background: #00ff88;")
        self.process_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.process_btn)

        # Progress
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Log
        layout.addWidget(QLabel("Live Log:"))
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMinimumHeight(220)
        layout.addWidget(self.log)

        self.setCentralWidget(central)

        self.json_files = []
        self.output_path = None

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select Keyboard Activity JSON files", "",
            "JSON Files (*.json);;All Files (*)"
        )
        if files:
            self.json_files = files
            self.file_list.clear()
            for f in files:
                self.file_list.addItem(f.split("/")[-1])
            self.log.append(f"✅ Selected {len(files)} JSON file(s)")

    def select_output(self):
        file, _ = QFileDialog.getSaveFileName(
            self, "Save Excel File", "Keyboard_Activity_Report.xlsx",
            "Excel Files (*.xlsx)"
        )
        if file:
            if not file.endswith(".xlsx"):
                file += ".xlsx"
            self.output_path = file
            self.output_label.setText(f"📍 {file}")

    def start_conversion(self):
        if not self.json_files:
            QMessageBox.warning(self, "No files", "Please select at least one JSON file")
            return
        if not self.output_path:
            QMessageBox.warning(self, "No output", "Please choose where to save the Excel file")
            return

        self.process_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.log.clear()
        self.log.append("🔥 Starting multi-user conversion...")

        self.thread = ConverterThread(self.json_files, self.output_path)
        self.thread.log_signal.connect(self.log.append)
        self.thread.progress_signal.connect(self.progress.setValue)
        self.thread.finished_signal.connect(self.conversion_finished)
        self.thread.start()

    def conversion_finished(self, success, message):
        self.process_btn.setEnabled(True)
        self.progress.setVisible(False)
        
        if success:
            QMessageBox.information(self, "🎉 Success!", message)
            self.log.append(f"{message}")
        else:
            QMessageBox.critical(self, "Error", message)
            self.log.append(f"ERROR: {message}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
                