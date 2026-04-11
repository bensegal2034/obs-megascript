import obsws_python as obs
from random import choice
from playsound import playsound
import ctypes as ct
from pathlib import Path
from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer
from random import choice
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
        self.banned_strings = [
            "windows default lock screen",
            "windows.ui.core.corewindow",
            "lockapp.exe",
            "xamlexplorerhostislandwindow",
            "explorer.exe",
            "mozilla firefox",
            "mozillawindowclass",
            "firefox.exe",
            "obs",
            "steam",
            "onecommander",
            "premiere",
            "photoshop",
            "terminal",
            "cmd",
            "visual studio",
            "notepad",
            "github",
            "mpv",
            "windows input experience",
            "program manager"
        ]
        self.emote_gen = self.get_emote()

        self.user32 = ct.windll.user32
        self.user32.SetProcessDPIAware()

        self.connect_attempts_interval = 5

        self.switcher_thread = None
        self.switcher_poll_interval = 1
        self.switcher_active = True

        self.change_tabbed_text_thread = None
        self.change_tabbed_text_poll_interval = 2

        self.buffer_timeout = 300
        self.afk_timer = 0

        # solution for getting script running dir stolen from here:
        # https://stackoverflow.com/a/9350788
        self.script_path = Path(os.path.dirname(os.path.realpath(__file__)))

        self.modified_times = {}

        self.connected = False
        self.running = True

        logging.basicConfig(
            filename=Path.joinpath(self.script_path, "megascript.log"),
            filemode="a",
            format='%(asctime)s %(levelname)s %(module)s - %(funcName)s: %(message)s',
            datefmt='%Y-%m-%d %I:%M:%S %p',
            level=logging.INFO
        )
        logging.getLogger("obsws_python").setLevel(logging.CRITICAL)
        self.logger = logging.getLogger("obs-megascript")
        self.logger_last_msg = ""

        self.log_info_norepeat("Script started, connecting to OBS...")
        self.establish_connection()
        self.log_info_norepeat("Connected to OBS!")

        self.commands_observer = None
        self.commands_event = None

    def log_info_norepeat(self, msg):
        if not msg == self.logger_last_msg:
            self.logger.info(msg)
        self.logger_last_msg = msg

    def establish_connection(self):
        while not self.connected:
            try:
                self.req = obs.ReqClient()
                self.evt = obs.EventClient()

                # run some test funcs to make sure we actually have a connection to obs
                self.req.get_version()
                self.req.get_current_program_scene()
                self.req.get_stats()

                # at this point if we haven't thrown an error we're probably chill
                # set up the event callbacks and set connected to true
                self.evt.callback.register(self.on_replay_buffer_saved)
                #self.evt.callback.register(self.on_replay_buffer_state_changed)
                self.evt.callback.register(self.on_record_state_changed)
                self.connected = True
                self.log_info_norepeat("Reached end of establish connection loop..")
            except Exception as error:
                self.log_info_norepeat(f"Errored out of loop with error: {error}")
                time.sleep(self.connect_attempts_interval)

    def handle_connection_lost(self, error):
        self.logger.exception(error)
        if self.connected: # only trigger this once so we don't have multiple instances of establish_connection() running
            self.connected = False
            self.running = False

            # kill all other threads
            if self.commands_observer:
                self.commands_observer.stop()

            ### AI WRITTEN EXPLANATION FOR THIS CODE
            # The issue arises because handle_connection_lost is called from within one of the threads (e.g., change_tabbed_text), 
            # and it attempts to join the same thread that's executing the method, which is not allowed.

            # To fix this, I've added checks to prevent joining the current thread by comparing the thread's ident with threading.current_thread().ident.

            # This ensures that when a thread calls handle_connection_lost, it won't try to join itself. 
            # The threads will exit their loops naturally when self.running is set to False, and new threads will be started after reconnection.
            if self.change_tabbed_text_thread and self.change_tabbed_text_thread.ident != threading.current_thread().ident:
                self.change_tabbed_text_thread.join(timeout=5)
                self.change_tabbed_text_thread = None
            if self.switcher_thread and self.switcher_thread.ident != threading.current_thread().ident:
                self.switcher_thread.join(timeout=5)
                self.switcher_thread = None

            self.logger.error(f"OBS connection failed, reconnecting...")
            self.establish_connection()
            
            # restart threads now that we're back online
            self.running = True
            if self.change_tabbed_text_thread is None:
                self.change_tabbed_text_thread = threading.Thread(target=self.change_tabbed_text)
                self.change_tabbed_text_thread.start()
            if self.switcher_thread is None:
                self.switcher_thread = threading.Thread(target=self.switcher)
                self.switcher_thread.start()

            self.log_info_norepeat(f"Reconnected to OBS!")

    def on_record_state_changed(self, data):
        saved_recording_data = None
        output_state = data.output_state
        if data.output_path is not None:
            saved_recording_data = Path(data.output_path)
        if output_state == "OBS_WEBSOCKET_OUTPUT_STOPPING":
            playsound(str(Path.joinpath(self.script_path, "recordingstartbeep.mp3")))
        elif output_state == "OBS_WEBSOCKET_OUTPUT_STOPPED":
            self.handle_saved_file(saved_recording_data)

    def on_replay_buffer_saved(self, data):
        playsound(str(Path.joinpath(self.script_path, "recordingstartbeep.mp3")))
        saved_replay = Path(data.saved_replay_path)
        self.handle_saved_file(saved_replay)

    # def on_replay_buffer_state_changed(self, data):
    #     output_state = data.output_state
    #     print(output_state)
    #     if output_state == "OBS_WEBSOCKET_OUTPUT_STOPPING":
    #         playsound(str(Path.joinpath(self.script_path, "recordingstartbeep.mp3")))

    def check_names_against_dir(self, names, dir):
        # this func stolen from here: https://stackoverflow.com/a/5320179
        def findWholeWord(w):
            return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search   
             
        valid_dir = None
        winning_name = None
        
        for root, dirs, files in dir.walk():
            for dir in dirs:
                for name in names:
                    if findWholeWord(dir)(name):
                        valid_dir = Path.joinpath(root, dir)
        
        # we still haven't found a dir if this triggers
        # iterate over dirs again, but this time match using in keyword
        if valid_dir is None:
            for root, dirs, files in dir.walk():
                for dir in dirs:
                    for name in names:
                        if dir.lower() in name.lower():
                            valid_dir = Path.joinpath(root, dir)
        
        return valid_dir, winning_name
    
    def handle_saved_file(self, filepath):
        recording_dir = filepath.parents[0]
        error = False
        moved = False
        fullscreen_windows = self.get_fullscreen_windows()
        names = []

        try:
            if fullscreen_windows:
                # check against all open fullscreen windows
                # technically this means that whatever window was addded earlier gets priority
                # however i don't know of a better way to do this atm so it is what it is
                # todo: maybe improve this

                for window_dict in fullscreen_windows.values():
                    names.append(window_dict["obs_window_str"])
                
                correct_dir, winning_name = self.check_names_against_dir(names, recording_dir)
                
                # make our own dir if none is found
                # use window name from first entry in fullscreen windows
                if correct_dir is None:
                    first_window_name = next(iter(fullscreen_windows.keys()))
                    correct_dir = Path.joinpath(recording_dir, first_window_name)
                    correct_dir.mkdir()

                shutil.move(filepath, correct_dir)
                moved = True
            else:
                error = True
                self.logger.error(f"Error moving '{filepath}'. No application was detected as being in fullscreen.")

        except Exception as error:
            error = True
            self.logger.exception(error)
            if moved:
                self.logger.error(f"Error moving '{filepath}'. File was moved from original location to '{correct_dir}'.")
            else:
                self.logger.error(f"Error moving '{filepath}'. File was NOT moved from original location. \nVars dump: 'names' = {names}, 'correct_dir' = {correct_dir}.")

        if not error:
            self.log_info_norepeat(f"Succesfully saved original file '{filepath}' at '{correct_dir}'. Valid names considered: '{names}'.")
            playsound(str(Path.joinpath(self.script_path, "recordingendbeep.mp3")))
        else:
            playsound(str(Path.joinpath(self.script_path, "recordingerror.mp3")))
    
    def get_fullscreen_windows(self):

        def is_hWnd_fullscreen(rect, full_screen_rect):
            rect_size_x = rect[2]
            rect_size_y = rect[3]
            fsr_size_x = full_screen_rect[2]
            fsr_size_y = full_screen_rect[3]
            fullscreen = False
            
            if rect_size_x >= fsr_size_x and rect_size_y >= fsr_size_y:
                fullscreen = True

            return fullscreen

        def win_enum_handler(hWnd, fullscreen_windows):
            full_screen_rect = fullscreen_windows["full_screen_rect"]
            rect = win32gui.GetWindowRect(hWnd)
            window_name = win32gui.GetWindowText(hWnd)
            if win32gui.IsWindowVisible(hWnd):
                if is_hWnd_fullscreen(rect, full_screen_rect):
                    if window_name != "":
                        fullscreen_window_dict = {}
                        tid, pid = win32process.GetWindowThreadProcessId(hWnd) # first var is thread id, second var is process id
                        proc = psutil.Process(pid)
                        exe_name = Path(proc.exe()).stem + ".exe"
                        class_name = win32gui.GetClassName(hWnd)
                        obs_window_str = f"{window_name}:{class_name}:{exe_name}"
                        window_str_safe = True

                        for banned_str in self.banned_strings:
                            if banned_str.lower() in obs_window_str.lower():
                                window_str_safe = False

                        if window_str_safe:
                            fullscreen_window_dict.update({
                                "hWnd": hWnd,
                                "tid": tid,
                                "pid": pid,
                                "proc": proc,
                                "exe_name": exe_name,
                                "class_name": class_name,
                                "obs_window_str": obs_window_str
                            })
                            fullscreen_windows[window_name] = fullscreen_window_dict

        try:
            fullscreen_windows = {
                "full_screen_rect": (0, 0, self.user32.GetSystemMetrics(0), self.user32.GetSystemMetrics(1))
            }
            win32gui.EnumWindows(win_enum_handler, fullscreen_windows)
            # remove this because no other functions really need it and it was a massive headache to deal with otherwise
            fullscreen_windows.pop("full_screen_rect")
        except Exception as error:
            self.logger.error(error)
            return False
        
        return fullscreen_windows

    def switcher(self):
        interval = self.switcher_poll_interval

        while self.running:
            if not self.switcher_active:
                time.sleep(interval)
                continue

            try:
                current_scene = self.req.get_current_program_scene().scene_name

                fullscreen_windows = self.get_fullscreen_windows()
                chosen_window_dict = None
                if len(fullscreen_windows) == 0:
                    self.log_info_norepeat(f"No fullscreen windows detected!")
                    is_in_foreground = False
                elif len(fullscreen_windows) == 1:
                    # get the first value from the dict
                    chosen_window_dict = next(iter(fullscreen_windows.values()))
                    is_in_foreground = chosen_window_dict["hWnd"] == self.user32.GetForegroundWindow()
                else:
                    self.log_info_norepeat(f"Multiple valid choices detected in fullscreen_windows with value: {fullscreen_windows}")
                    # multiple valid choices to switch to as more than 1 fullscreen window has been detected
                    # check if any of them match our recording directory, if so, use that one
                    # otherwise, fuck it dude just pick at random
                    # todo: improve this
                    names = []
                    recording_directory = self.req.get_record_directory().record_directory
                    for window_dict in fullscreen_windows.values():
                        names.append(window_dict["obs_window_str"])
                    valid_dir, winning_name = self.check_names_against_dir(names, recording_directory)
                    if valid_dir:
                        for window_dict in fullscreen_windows.values():
                            if window_dict["obs_window_str"] == winning_name:
                                chosen_window_dict = window_dict
                    else:
                        chosen_window_dict = choice(list(fullscreen_windows.values()))
                    self.log_info_norepeat(f"Chose {chosen_window_dict} from multiple choice fullscreen_windows dict")
                    is_in_foreground = chosen_window_dict["hWnd"] == self.user32.GetForegroundWindow()
            
            except Exception as error:
                self.handle_connection_lost(error)

            if not is_in_foreground:
                try:
                    if current_scene != "Alt Tabbed":
                        self.log_info_norepeat("Setting scene to Alt Tabbed.")
                        self.req.set_current_program_scene("Alt Tabbed")
                        self.afk_timer = int(time.time()) + self.buffer_timeout
                except Exception as error:
                    self.handle_connection_lost(error)
            else:
                try:
                    if current_scene != "Game Capture":
                        self.log_info_norepeat(f"Setting scene to Game Capture, switching Game Capture output to {chosen_window_dict["obs_window_str"]}.")
                        self.req.set_current_program_scene("Game Capture")
                        self.req.set_input_settings(
                            name="Capture 0", 
                            settings={
                                "capture_mode": "window",
                                "window": chosen_window_dict["obs_window_str"]
                            },
                            overlay=True
                        )
                        
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
                self.log_info_norepeat(f"Stopping replay buffer, current time '{now}' greater than afk timer '{self.afk_timer}' and replay buffer active.")
                self.req.stop_replay_buffer()
            
            elif current_scene == "Game Capture" and not(buffer_active):
                self.log_info_norepeat("Starting replay buffer.")
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

        while self.running:
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
            self.change_tabbed_text_thread = threading.Thread(target=self.change_tabbed_text)
            self.change_tabbed_text_thread.start()
        
        # start the switcher
        if self.switcher_thread is None:
            self.switcher_thread = threading.Thread(target=self.switcher)
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
                                self.log_info_norepeat(f"Toggling switcher to {self.switcher_active}.")
                                commands_data["toggleSwitcher"] = False
                            
                        with open(commands_path, "w") as f:
                            json.dump(commands_data, f)

        # start observer thread for commands
        self.commands_event = CommandsEvent()
        self.commands_observer = Observer()
        self.commands_observer.schedule(
            event_handler = self.commands_event, 
            path = self.script_path,
            recursive = False
        )
        self.commands_observer.start()

        # keep the main thread alive
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.logger.info("Received KeyboardInterrupt, shutting down...")
            self.running = False
            self.commands_observer.stop()
            self.commands_observer.join()
            if self.change_tabbed_text_thread:
                self.change_tabbed_text_thread.join(timeout=5)
            if self.switcher_thread:
                self.switcher_thread.join(timeout=5)
            self.logger.info("Shutdown complete.")

if __name__ == "__main__":
    ms = MegaScript()
    ms.run()