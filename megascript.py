import obsws_python as obs
from random import choice
from playsound import playsound
import ctypes as ct
from pathlib import Path
import win32gui, win32process, time, threading, psutil, shutil

class MegaScript:

    def __init__(self):
        self.evt = None
        self.req = None
        self.textUpdateDelay = 2
        self.emoticons = [
            ":3",
            ":)",
            ":D",
            ":P",
            ":>",
            "xd",
            "uwu",
            "owo",
            ":O",
            ":3c",
            "3:",
            "c:"
        ]
        self.emote_gen = self.get_emote()

        self.obs_connection_timeout = 30
        self.req = obs.ReqClient(host='localhost', port=4455, password='arcane2455', timeout=self.obs_connection_timeout)
        self.evt = obs.EventClient(host='localhost', port=4455, password='arcane2455', timeout=self.obs_connection_timeout)
        self.evt.callback.register(self.on_replay_buffer_saved)

        self.user32 = ct.windll.user32
        self.user32.SetProcessDPIAware()

        self.switcher_thread = None
        self.switcher_poll_interval = 1

        self.change_tabbed_text_thread = None
        self.change_tabbed_text_poll_interval = 2

        self.buffer_timeout = 300
        self.afk_timer = 0

    def on_replay_buffer_saved(self, data):
        saved_replay = Path(data.saved_replay_path)
        recording_dir = saved_replay.parents[0]
        error = False

        try:
            if self.is_fullscreen():
                hWnd = self.user32.GetForegroundWindow()
                tid, pid = win32process.GetWindowThreadProcessId(hWnd) # first var is thread id, second var is process id
                proc = psutil.Process(pid)
                names = [Path(proc.exe()).stem, proc.name(), win32gui.GetWindowText(hWnd)]
                correct_dir = None
                
                for root, dirs, files in recording_dir.walk():
                    for dir in dirs:
                        for name in names:
                            if dir.lower() in name.lower():
                                correct_dir = Path.joinpath(root, dir)
                
                if correct_dir is None:
                    # make our own dir using the exe name if no valid dir is found
                    correct_dir = Path.joinpath(recording_dir, names[0])
                    correct_dir.mkdir()

                shutil.move(saved_replay, correct_dir)
            else:
                print("Not fullscreen")
                error = True
        except Exception as e:
            print(e)
            error = True

        if not error:
            playsound("D:\\Music\\recordingbeep.mp3")
        else:
            playsound("D:\\Music\\recordingerror.mp3")
    
    def is_fullscreen(self):
        full_screen_rect = (0, 0, self.user32.GetSystemMetrics(0), self.user32.GetSystemMetrics(1))
        try:
            hWnd = self.user32.GetForegroundWindow()
            rect = win32gui.GetWindowRect(hWnd)
            
            return rect == full_screen_rect
        except:
            return False

    def switcher(self):
        interval = self.switcher_poll_interval

        while True:
            if not(self.is_fullscreen()):
                current_scene = self.req.get_current_program_scene().scene_name
                try:
                    if current_scene != "Alt Tabbed":
                        self.req.set_current_program_scene("Alt Tabbed")
                        self.afk_timer = int(time.time()) + self.buffer_timeout
                except:
                    pass
            else:
                try:
                    if current_scene != "Game Capture":
                        self.req.set_current_program_scene("Game Capture")
                except:
                    pass
            
            self.manage_buffer_state()
            
            time.sleep(interval)
            continue

    def manage_buffer_state(self):
        try:
            current_scene = self.req.get_current_program_scene().scene_name
            buffer_active = self.req.get_replay_buffer_status().output_active

            if current_scene == "Alt Tabbed" and int(time.time()) >= self.afk_timer and buffer_active:
                self.req.stop_replay_buffer()
            
            elif current_scene == "Game Capture" and not(buffer_active):
                self.req.start_replay_buffer()
        except:
            pass

    def get_emote(self):
        previous_emote = None
        while True:
            emote = choice(self.emoticons)
            if emote != previous_emote:
                yield emote
                previous_emote = emote
    
    def change_tabbed_text(self):
        interval = self.change_tabbed_text_poll_interval

        while True:
            try:
                if self.req.get_current_program_scene().scene_name == "Alt Tabbed":
                    self.req.set_input_settings("Alt Tabbed Text", {
                        "text": f"Alt Tabbed {next(self.emote_gen)}"
                    }, True)
            except:
                pass

            time.sleep(interval)
            continue

    def run(self):
        # start the text changer
        if self.change_tabbed_text_thread is None:
            self.change_tabbed_text_thread = threading.Thread(target=self.change_tabbed_text, daemon=True)
            self.change_tabbed_text_thread.start()
        
        # start the switcher
        if self.switcher_thread is None:
            self.switcher_thread = threading.Thread(target=self.switcher, daemon=True)
            self.switcher_thread.start()

        # non blocking; keep main thread alive so the other children threads can do their jobs
        keepalive = threading.Event()
        keepalive.wait()

if __name__ == "__main__":
    ms = MegaScript()
    ms.run()