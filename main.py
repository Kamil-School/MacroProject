import sys
import time
import json
import os
from PyQt5 import QtWidgets, QtCore, QtGui
from pynput import keyboard, mouse

def serialize_key(key):
    try:
        if hasattr(key, 'char') and key.char is not None:
            return {"vtype": "KeyCode", "char": key.char}
        else:
            return {"vtype": "Key", "name": key.name}
    except AttributeError:
        return {"vtype": "str", "value": str(key)}

def deserialize_key(data):
    if data["vtype"] == "KeyCode":
        return keyboard.KeyCode.from_char(data["char"])
    elif data["vtype"] == "Key":
        return getattr(keyboard.Key, data["name"])
    else:
        return keyboard.KeyCode.from_char(data.get("value", ""))

def serialize_mouse_button(button):
    return button.name

def deserialize_mouse_button(name):
    return getattr(mouse.Button, name)

def serialize_event(event):
    event_time, event_type, event_data = event
    result = {"time": event_time, "type": event_type}
    if event_type in ["key_press", "key_release"]:
        result["data"] = serialize_key(event_data)
    elif event_type == "mouse_move":
        result["data"] = {"x": event_data[0], "y": event_data[1]}
    elif event_type == "mouse_click":
        x, y, button, pressed = event_data
        result["data"] = {
            "x": x,
            "y": y,
            "button": serialize_mouse_button(button),
            "pressed": pressed
        }
    elif event_type == "mouse_scroll":
        x, y, dx, dy = event_data
        result["data"] = {"x": x, "y": y, "dx": dx, "dy": dy}
    else:
        result["data"] = event_data
    return result

def deserialize_event(event_dict):
    event_time = event_dict["time"]
    event_type = event_dict["type"]
    data = event_dict["data"]
    if event_type in ["key_press", "key_release"]:
        key = deserialize_key(data)
        return event_time, event_type, key
    elif event_type == "mouse_move":
        return event_time, event_type, (data["x"], data["y"])
    elif event_type == "mouse_click":
        button = deserialize_mouse_button(data["button"])
        return event_time, event_type, (data["x"], data["y"], button, data["pressed"])
    elif event_type == "mouse_scroll":
        return event_time, event_type, (data["x"], data["y"], data["dx"], data["dy"])
    else:
        return event_time, event_type, data

class MacroRecorder:
    def __init__(self):
        self.events = []
        self.recording = False
        self.start_time = None
        self.stop_hotkey = keyboard.Key.f12
        self.hotkey_triggered = False
        self.on_stop = None

    def start_recording(self):
        self.events = []
        self.recording = True
        self.start_time = time.time()
        self.hotkey_triggered = False

        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_key_press,
            on_release=self.on_key_release
        )
        self.mouse_listener = mouse.Listener(
            on_click=self.on_click,
            on_scroll=self.on_scroll
        )
        self.keyboard_listener.start()
        self.mouse_listener.start()

    def stop_recording(self):
        if not self.recording:
            return self.events
        self.recording = False
        self.keyboard_listener.stop()
        self.mouse_listener.stop()
        if self.on_stop:
            QtCore.QTimer.singleShot(0, lambda: self.on_stop(self.events))
        return self.events

    def stop_recording_from_hotkey(self):
        if self.recording:
            self.stop_recording()

    def record_event(self, event_type, event_data):
        if not self.recording:
            return
        event_time = time.time() - self.start_time
        self.events.append((event_time, event_type, event_data))

    def on_key_press(self, key):
        if key == self.stop_hotkey:
            if not self.hotkey_triggered:
                self.hotkey_triggered = True
                QtCore.QTimer.singleShot(0, self.stop_recording_from_hotkey)
            return
        self.record_event('key_press', key)

    def on_key_release(self, key):
        if key == self.stop_hotkey:
            return
        self.record_event('key_release', key)

    def on_click(self, x, y, button, pressed):
        self.record_event('mouse_click', (x, y, button, pressed))

    def on_scroll(self, x, y, dx, dy):
        self.record_event('mouse_scroll', (x, y, dx, dy))

class MacroPlayer:
    def __init__(self, events):
        self.events = events
        self._stop = False

    def play(self):
        kb_controller = keyboard.Controller()
        mouse_controller = mouse.Controller()
        start_time = time.time()
        for event in self.events:
            if self._stop:
                break

            event_time, event_type, event_data = event
            time_to_wait = event_time - (time.time() - start_time)
            if time_to_wait > 0:
                time.sleep(time_to_wait)
            if self._stop:
                break

            if event_type == 'key_press':
                try:
                    kb_controller.press(event_data)
                except Exception as e:
                    print(f"Error on key press: {e}")
            elif event_type == 'key_release':
                try:
                    kb_controller.release(event_data)
                except Exception as e:
                    print(f"Error on key release: {e}")
            elif event_type == 'mouse_click':
                x, y, button, pressed = event_data
                if pressed:
                    mouse_controller.press(button)
                else:
                    mouse_controller.release(button)
            elif event_type == 'mouse_scroll':
                x, y, dx, dy = event_data
                mouse_controller.scroll(dx, dy)

class MacroPlayerWorker(QtCore.QObject):
    finished = QtCore.pyqtSignal()

    def __init__(self, events):
        super().__init__()
        self.events = events
        self.player = None

    @QtCore.pyqtSlot()
    def run(self):
        self.player = MacroPlayer(self.events)
        self.player.play()
        self.finished.emit()

    @QtCore.pyqtSlot()
    def stop(self):
        if self.player is not None:
            self.player._stop = True

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Macro Program")
        self.setGeometry(100, 100, 800, 600)
        self.recorder = MacroRecorder()

        self.macros_folder = "macros"
        os.makedirs(self.macros_folder, exist_ok=True)

        self.initUI()
        self.initShortcuts()

    def initUI(self):
        self.tabs = QtWidgets.QTabWidget()
        self.setCentralWidget(self.tabs)

        self.record_tab = QtWidgets.QWidget()
        self.tabs.addTab(self.record_tab, "Record Macro")
        self.record_layout = QtWidgets.QVBoxLayout()
        self.record_tab.setLayout(self.record_layout)

        self.start_button = QtWidgets.QPushButton("Start Recording")
        self.start_button.clicked.connect(self.start_recording)
        self.stop_button = QtWidgets.QPushButton("Stop Recording")
        self.stop_button.clicked.connect(self.stop_recording)
        self.stop_button.setEnabled(False)
        self.record_layout.addWidget(self.start_button)
        self.record_layout.addWidget(self.stop_button)

        self.countdown_label = QtWidgets.QLabel("")
        self.record_layout.addWidget(self.countdown_label)

        self.event_list = QtWidgets.QListWidget()
        self.record_layout.addWidget(self.event_list)

        btn_layout = QtWidgets.QHBoxLayout()
        self.save_button = QtWidgets.QPushButton("Save Macro")
        self.save_button.clicked.connect(self.save_macro)
        btn_layout.addWidget(self.save_button)
        self.record_layout.addLayout(btn_layout)

        self.play_tab = QtWidgets.QWidget()
        self.tabs.addTab(self.play_tab, "Play Macro")
        self.play_layout = QtWidgets.QVBoxLayout()
        self.play_tab.setLayout(self.play_layout)

        macro_select_layout = QtWidgets.QHBoxLayout()
        macro_select_layout.addWidget(QtWidgets.QLabel("Select Macro:"))
        self.macro_dropdown = QtWidgets.QComboBox()
        macro_select_layout.addWidget(self.macro_dropdown)
        self.refresh_button = QtWidgets.QPushButton("Refresh List")
        self.refresh_button.clicked.connect(self.refresh_macro_list)
        macro_select_layout.addWidget(self.refresh_button)
        self.play_layout.addLayout(macro_select_layout)
        self.refresh_macro_list()

        self.play_button = QtWidgets.QPushButton("Play Macro")
        self.play_button.clicked.connect(self.play_macro)
        self.play_layout.addWidget(self.play_button)

        self.stop_play_button = QtWidgets.QPushButton("Stop Macro")
        self.stop_play_button.clicked.connect(self.stop_macro)
        self.stop_play_button.setEnabled(False)
        self.play_layout.addWidget(self.stop_play_button)

        self.play_countdown_label = QtWidgets.QLabel("")
        self.play_layout.addWidget(self.play_countdown_label)

    def initShortcuts(self):
        self.shortcut_start_playback = QtWidgets.QShortcut(QtGui.QKeySequence("F9"), self)
        self.shortcut_start_playback.activated.connect(self.play_macro)

        self.shortcut_stop_playback = QtWidgets.QShortcut(QtGui.QKeySequence("F8"), self)
        self.shortcut_stop_playback.activated.connect(self.stop_macro)

    def refresh_macro_list(self):
        self.macro_dropdown.clear()
        macro_files = [f for f in os.listdir(self.macros_folder) if f.endswith(".json")]
        if not macro_files:
            self.macro_dropdown.addItem("No macros found")
            self.macro_dropdown.setEnabled(False)
        else:
            self.macro_dropdown.addItems(macro_files)
            self.macro_dropdown.setEnabled(True)

    def start_recording(self):
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(False)
        self.event_list.clear()
        self.countdown_value = 3
        self.countdown_label.setText(f"Recording starts in {self.countdown_value} seconds...")
        self.countdown_timer = QtCore.QTimer()
        self.countdown_timer.timeout.connect(self.update_countdown)
        self.countdown_timer.start(1000)

    def update_countdown(self):
        self.countdown_value -= 1
        if self.countdown_value > 0:
            self.countdown_label.setText(f"Recording starts in {self.countdown_value} seconds...")
        else:
            self.countdown_timer.stop()
            self.countdown_label.setText("")
            self.begin_recording()

    def begin_recording(self):
        self.recorder.on_stop = self.on_recording_stopped
        self.recorder.start_recording()
        self.stop_button.setEnabled(True)
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.update_event_list)
        self.timer.start(100)
        self.event_list.addItem("Recording started")
        self.event_list.addItem("Press F12 to stop recording.")

    def stop_recording(self):
        if not self.recorder.recording:
            return
        if hasattr(self, 'timer'):
            self.timer.stop()
        events = self.recorder.stop_recording()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.event_list.addItem(f"Recording stopped. Total events: {len(events)}")

    def on_recording_stopped(self, events):
        if hasattr(self, 'timer'):
            self.timer.stop()
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.event_list.addItem(f"Recording stopped (hotkey). Total events: {len(events)}")

    def update_event_list(self):
        count = len(self.recorder.events)
        self.event_list.clear()
        self.event_list.addItem(f"Recording... Events recorded: {count}")
        self.event_list.addItem("Press F12 to stop recording.")

    def save_macro(self):
        if not self.recorder.events:
            QtWidgets.QMessageBox.warning(self, "No Macro", "There is no macro to save")
            return

        name, ok = QtWidgets.QInputDialog.getText(self, "Save Macro", "Enter macro name:")
        if ok and name:
            filename = os.path.join(self.macros_folder, f"{name}.json")
            try:
                with open(filename, "w") as outfile:
                    json_events = [serialize_event(e) for e in self.recorder.events]
                    json.dump(json_events, outfile, indent=4)
                self.event_list.addItem(f"Macro saved as {filename}")
                self.refresh_macro_list()
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to save macro: {e}")

    def play_macro(self):
        macro_file = self.macro_dropdown.currentText()
        if not macro_file or macro_file == "No macros found":
            QtWidgets.QMessageBox.warning(self, "No Macro", "No macro selected")
            return

        filename = os.path.join(self.macros_folder, macro_file)
        try:
            with open(filename, "r") as infile:
                json_events = json.load(infile)
            self.events_to_play = [deserialize_event(e) for e in json_events]
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load macro: {e}")
            return

        self.event_list.clear()
        self.event_list.addItem("Preparing to play macro...")
        self.play_button.setEnabled(False)

        self.play_countdown_value = 3
        self.play_countdown_label.setText(f"Playback starts in {self.play_countdown_value} seconds...")
        self.play_countdown_timer = QtCore.QTimer()
        self.play_countdown_timer.timeout.connect(self.update_play_countdown)
        self.play_countdown_timer.start(1000)

    def update_play_countdown(self):
        self.play_countdown_value -= 1
        if self.play_countdown_value > 0:
            self.play_countdown_label.setText(f"Playback starts in {self.play_countdown_value} seconds...")
        else:
            self.play_countdown_timer.stop()
            self.play_countdown_label.setText("")
            self.begin_playback(self.events_to_play)

    def begin_playback(self, events):
        self.event_list.addItem("Playing macro...")
        self.thread = QtCore.QThread()
        self.worker = MacroPlayerWorker(events)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_macro_play_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()
        self.stop_play_button.setEnabled(True)

    def stop_macro(self):
        if hasattr(self, 'play_countdown_timer') and self.play_countdown_timer.isActive():
            self.play_countdown_timer.stop()
            self.play_countdown_label.setText("")
            self.event_list.addItem("Playback canceled")
            self.play_button.setEnabled(True)
            return

        if hasattr(self, 'worker'):
            self.worker.stop()
            self.event_list.addItem("Macro playback stopped")
            self.stop_play_button.setEnabled(False)

    def on_macro_play_finished(self):
        self.event_list.addItem("Macro playback finished")
        self.play_button.setEnabled(True)
        self.stop_play_button.setEnabled(False)

def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
