import obsws_python as obs
from random import choice
from playsound import playsound
import ctypes as ct
from pathlib import Path
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
import win32gui, win32process, time, threading, psutil, shutil, logging, re, os, json

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

        self.user32 = ct.windll.user32
        self.user32.SetProcessDPIAware()

        self.switcher_thread = None
        self.switcher_poll_interval = 1
        self.switcher_active = True

        self.change_tabbed_text_thread = None
        self.change_tabbed_text_poll_interval = 2

        self.buffer_timeout = 300
        self.afk_timer = 0

        # absolute path sucks but is annoying to fix
        self.script_path = Path("C:\\Users\\Ben\\Scripts\\obs-megascript")

        self.modified_times = {}

        self.connected = False

        logging.basicConfig(
            filename=Path.joinpath(self.script_path, "megascript.log"),
            filemode="a",
            format='%(asctime)s %(levelname)s %(module)s - %(funcName)s: %(message)s',
            datefmt='%Y-%m-%d %I:%M:%S %p',
            level=logging.INFO
        )
        self.logger = logging.getLogger(__name__)

        self.logger.info("Script started, connecting to OBS...")
        self.establish_connection()
        self.logger.info("Connected to OBS!")

    def establish_connection(self):
        self.obs_connection_timeout = 5
        self.connect_attempts_interval = 5

        while not self.connected:
            try:
                self.req = obs.ReqClient(host='localhost', port=4455, password='arcane2455', timeout=self.obs_connection_timeout)
                self.evt = obs.EventClient(host='localhost', port=4455, password='arcane2455', timeout=self.obs_connection_timeout)

                # run some test funcs to make sure we actually have a connection to obs
                self.req.get_version()
                self.req.get_current_program_scene()
                self.req.get_stats()

                # at this point if we haven't thrown an error we're probably chill
                # set up the event callbacks and set connected to true
                self.evt.callback.register(self.on_replay_buffer_saved)
                self.connected = True
            except:
                time.sleep(self.connect_attempts_interval)

    def handle_connection_lost(self, error):
        self.logger.exception(error)
        if self.connected: # only trigger this once so we don't have multiple instances of establish_connection() running
            self.connected = False

            self.logger.error(f"OBS connection failed, reconnecting...")
            self.establish_connection()
            self.logger.info(f"Reconnected to OBS!")

    def on_replay_buffer_saved(self, data):
        saved_replay = Path(data.saved_replay_path)
        recording_dir = saved_replay.parents[0]
        error = False
        moved = False

        # this func stolen from here: https://stackoverflow.com/a/5320179
        def findWholeWord(w):
            return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search

        try:
            if self.is_fullscreen():
                hWnd = self.user32.GetForegroundWindow()
                tid, pid = win32process.GetWindowThreadProcessId(hWnd) # first var is thread id, second var is process id
                proc = psutil.Process(pid)

                names = [Path(proc.exe()).stem, win32gui.GetWindowText(hWnd)] 
                correct_dir = None
                
                for root, dirs, files in recording_dir.walk():
                    for dir in dirs:
                        for name in names:
                            if findWholeWord(dir)(name):
                                correct_dir = Path.joinpath(root, dir)
                
                # we still haven't found a dir if this triggers
                # iterate over dirs again, but this time match using in keyword
                if correct_dir is None:
                    for root, dirs, files in recording_dir.walk():
                        for dir in dirs:
                            for name in names:
                                if dir.lower() in name.lower():
                                    correct_dir = Path.joinpath(root, dir)
                
                # still haven't found a correct dir, all search options exhausted
                # make our own dir using the exe name if no valid dir is found
                if correct_dir is None:
                    correct_dir = Path.joinpath(recording_dir, names[0])
                    correct_dir.mkdir()

                shutil.move(saved_replay, correct_dir)
                moved = True
            else:
                error = True
                self.logger.error(f"Error moving '{saved_replay}'. No application was detected as being in fullscreen.")

        except Exception as error:
            error = True
            self.logger.exception(error)
            if moved:
                self.logger.error(f"Error moving '{saved_replay}'. File was moved from original location to '{correct_dir}'.")
            else:
                self.logger.error(f"Error moving '{saved_replay}'. File was NOT moved from original location. \nVars dump: 'names' = {names}, 'correct_dir' = {correct_dir}.")

        if not error:
            self.logger.info(f"Succesfully saved original file '{saved_replay}' at '{correct_dir}'. Valid names considered: '{names}'.")
            playsound("D:\\Music\\recordingbeep.mp3")
        else:
            playsound("D:\\Music\\recordingerror.mp3")
    
    def is_fullscreen(self):
        full_screen_rect = (0, 0, self.user32.GetSystemMetrics(0), self.user32.GetSystemMetrics(1))
        try:
            hWnd = self.user32.GetForegroundWindow()
            rect = win32gui.GetWindowRect(hWnd)
            
            return rect == full_screen_rect
        except Exception as error:
            self.logger.exception(error)
            return False

    def switcher(self):
        interval = self.switcher_poll_interval

        while True:
            if not self.switcher_active:
                time.sleep(interval)
                continue
            current_scene = self.req.get_current_program_scene().scene_name
            if not(self.is_fullscreen()):
                try:
                    if current_scene != "Alt Tabbed":
                        self.logger.info("Setting scene to Alt Tabbed.")
                        self.req.set_current_program_scene("Alt Tabbed")
                        self.afk_timer = int(time.time()) + self.buffer_timeout
                except Exception as error:
                    self.handle_connection_lost(error)
            else:
                try:
                    if current_scene != "Game Capture":
                        self.logger.info("Setting scene to Game Capture.")
                        self.req.set_current_program_scene("Game Capture")
                except Exception as error:
                    self.handle_connection_lost(error)
            
            self.manage_buffer_state()
            
            time.sleep(interval)
            continue

    def manage_buffer_state(self):
        try:
            current_scene = self.req.get_current_program_scene().scene_name
            buffer_active = self.req.get_replay_buffer_status().output_active

            now = int(time.time())
            if current_scene == "Alt Tabbed" and now >= self.afk_timer and buffer_active:
                self.logger.info(f"Stopping replay buffer, current time '{now}' greater than afk timer '{self.afk_timer}' and replay buffer active.")
                self.req.stop_replay_buffer()
            
            elif current_scene == "Game Capture" and not(buffer_active):
                self.logger.info("Starting replay buffer.")
                self.req.start_replay_buffer()
        except Exception as error:
            self.handle_connection_lost(error)

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
            except Exception as error:
                self.handle_connection_lost(error)

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

        # create event handler for commands
        class CommandsEvent(FileSystemEventHandler):
            def __init__(self):
                super().__init__()
            
            @staticmethod # this needs to be here so that self references the superclass and not CommandsEvent
            def on_any_event(event):

                # this try except block prevents duplicate events from occurring
                # stolen from here: https://stackoverflow.com/a/79415551
                try:
                    t = os.path.getmtime(event.src_path)
                    if event.src_path in self.modified_times and t == self.modified_times[event.src_path]:
                        # duplicate event
                        return
                    self.modified_times[event.src_path] = t
                except FileNotFoundError:
                    # file got deleted after event was triggered
                    try:
                        del self.modified_times[event.src_path]
                    except KeyError:
                        pass
                # continue processing event

                if event.is_directory:
                    return None

                elif event.event_type == 'modified':
                    if "commands.json" in event.src_path:
                        commands_path = Path.joinpath(self.script_path, "commands.json")
                        commands_data = None

                        with open(commands_path, "r") as f:
                            commands_data = json.load(f)
                            if commands_data["toggleSwitcher"]:
                                self.switcher_active = not(self.switcher_active)
                                self.logger.info(f"Toggling switcher to {self.switcher_active}.")
                                commands_data["toggleSwitcher"] = False
                            
                        with open(commands_path, "w") as f:
                            json.dump(commands_data, f)

        # start observer thread for commands
        commandsEvent = CommandsEvent()
        commandsObserver = Observer()
        commandsObserver.schedule(
            event_handler = commandsEvent, 
            path = self.script_path,
            recursive = False
        )
        commandsObserver.start()

        # non blocking; keep main thread alive so the other children threads can do their jobs
        keepalive = threading.Event()
        keepalive.wait()

if __name__ == "__main__":
    ms = MegaScript()
    ms.run()