import obsws_python as obs
from threading import Timer, Event
from random import choice
from playsound import playsound
import ctypes as ct
import json, os, time, threading
import tkinter as tk
from tkinter import messagebox, simpledialog
import win32gui, win32process, wmi, pythoncom

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

        with open(os.path.join(os.path.dirname(__file__), "programs.json")) as f:
            self.programsCache = json.load(f)

        # path to notify.json (same directory as this script)
        self.notify_path = os.path.join(os.path.dirname(__file__), "notify.json")
        self.notify_poll_interval = 3
        self._notify_thread = None

        self.program_path = os.path.join(os.path.dirname(__file__), "programs.json")

        self._switcher_thread = None
        self.switcher_poll_interval = 1

        self._change_tabbed_text_thread = None
        self._change_tabbed_text_poll_interval = 2
            

    def check_connection(self):
        connected = False
        ready = False
        try:
            self.req.get_version()
        except:
            while not(connected):
                try:
                    self.req = obs.ReqClient(host='localhost', port=4455, password='arcane2455', timeout=5)
                    self.evt = obs.EventClient(host='localhost', port=4455, password='arcane2455', timeout=5)
                    self.evt.callback.register(self.on_replay_buffer_saved)
                except:
                    connected = False
                else:
                    connected = True
        
        while not(ready):
            try:
                self.req.get_current_program_scene()
            except:
                ready = False
            else:
                ready = True

    def on_replay_buffer_saved(self, data):
        self.check_connection()
        playsound("D:\\Music\\recordingbeep.mp3")

    def check_requested_add(self):
        self.check_connection()
        # some of this is ai generated
        interval = self.notify_poll_interval

        while True:
            try:
                if not os.path.exists(self.notify_path):
                    time.sleep(interval)
                    continue

                with open(self.notify_path, "r", encoding="utf-8") as f:
                    try:
                        notify_data = json.load(f)
                    except Exception:
                        # malformed JSON; skip until next poll
                        time.sleep(interval)
                        continue

                # Only act when triggeredAddProgram is present and True
                if notify_data.get("triggeredAddProgram") is True:
                    # set boolean to false in the notify.json so that we're ready for the next add
                    # regardless of what happens next
                    notify_data["triggeredAddProgram"] = False
                    with open(self.notify_path, "w", encoding="utf-8") as f:
                        json.dump(notify_data, f, indent=4)

                    # setup vars and stuff for later
                    pref_name = None
                    current_win = None
                    current_win_exe = None
                    root_simple_dialog = tk.Tk()
                    root_simple_dialog.withdraw()
                    root_simple_dialog.attributes("-topmost", True)
                    with open(self.program_path, "r", encoding="utf-8") as f:
                        try:
                            program_data = json.load(f)
                        except Exception:
                            # uh oh spaghettios
                            messagebox.showerror(
                                title="help me god",
                                message=f"programs.json is horribly fucked its fucking over",
                                parent=root_simple_dialog
                            )
                            time.sleep(interval)
                            continue
                    # get active window exe name
                    pythoncom.CoInitialize() # this line is required for wmi stuff to work inside a thread. i dont know what it does.
                    wmiObj = wmi.WMI()
                    pycwnd = win32gui.GetForegroundWindow()
                    tid, pid = win32process.GetWindowThreadProcessId(pycwnd)
                    for proc in wmiObj.Win32_Process():
                        if proc.ProcessId == pid:
                            current_win = win32gui.GetWindowText(pycwnd)
                            current_win_exe = proc.ExecutablePath
                    
                    if current_win == None:
                        messagebox.showerror(
                            title="Error",
                            message=f"Could not get exe name for active window!",
                            parent=root_simple_dialog
                        )
                        time.sleep(interval)
                        continue  

                    # askokcancel returns True for OK, False for Cancel
                    add_ok = messagebox.askokcancel(
                        title="Add program to OBS capture",
                        message=f"Ok to add: {current_win}?",
                        parent=root_simple_dialog,
                    )

                    if add_ok:
                        # ask the user for a preferred name for the current window
                        pref_name = simpledialog.askstring(
                            title="Preferred name",
                            prompt=f"Preferred name for {current_win}?",
                            parent=root_simple_dialog,
                        )

                        if pref_name:

                            # define scene to switch to in obs
                            sceneList = self.req.get_scene_list().scenes
                            sceneNames = []
                            for sceneDict in sceneList:
                                sceneNames.append(sceneDict.get("sceneName"))

                            if not sceneNames:
                                sceneNames = ['(No scenes found)']

                            selectedSceneName = None
                            selectedSceneUuid = None

                            # the below stuff is ai generated code that i refactored to actually work
                            rootScenePicker = tk.Tk()
                            rootScenePicker.title("Pick OBS scene")
                            rootScenePicker.attributes('-topmost', True)
                            scenePickerSelector = tk.StringVar(rootScenePicker, value=sceneNames[1])
                            def on_confirm():
                                nonlocal selectedSceneName # this makes it so that selectedScene is the variable we assigned earlier, and is not confined to the on_confirm func
                                selectedSceneName = scenePickerSelector.get()
                                # apparently both of these are necessary??? i dont know why but it works
                                rootScenePicker.destroy()
                                rootScenePicker.quit()
                            frm = tk.Frame(rootScenePicker, padx=12, pady=12)
                            frm.pack(fill='both', expand=True)
                            tk.Label(frm, text='Choose scene:').pack(anchor='w')
                            tk.OptionMenu(frm, scenePickerSelector, *sceneNames).pack(fill='x', pady=(6,8))
                            tk.Button(frm, text='Confirm', width=12, command=on_confirm).pack()
                            rootScenePicker.mainloop()

                            for sceneDict in sceneList:
                                if sceneDict.get("sceneName") == selectedSceneName:
                                    selectedSceneUuid = sceneDict.get("sceneUuid")
                                    break

                            # we have all the data we need, go ahead and add it to the program list 
                            program_data[current_win] = {
                                "exe": current_win_exe,
                                "pref_name": pref_name,
                                "sceneUuid": selectedSceneUuid
                            }

                            with open(self.program_path, "w", encoding="utf-8") as f:
                                json.dump(program_data, f, indent=4)

                    # destroy the dialog root now that we're done
                    try:
                        root_simple_dialog.destroy()
                    except Exception:
                        pass
                    pythoncom.CoUninitialize()
                    with open(os.path.join(os.path.dirname(__file__), "programs.json")) as f:
                        self.programsCache = json.load(f)
                    time.sleep(interval)
                    continue
                    

                time.sleep(interval)
            except Exception:
                # keep the monitor alive even on unexpected errors
                time.sleep(interval)
    
    def is_fullscreen(self):
        full_screen_rect = (0, 0, self.user32.GetSystemMetrics(0), self.user32.GetSystemMetrics(1))
        try:
            hWnd = self.user32.GetForegroundWindow()
            rect = win32gui.GetWindowRect(hWnd)
            
            return [rect == full_screen_rect, hWnd]
        except:
            return [False, None]

    def switcher(self):
        interval = self.switcher_poll_interval

        while True:
            self.check_connection()
            is_fullscreen = self.is_fullscreen()

            if not is_fullscreen[0]:
                try:
                    self.req.set_current_program_scene("Alt Tabbed")
                except:
                    pass
            else:
                try:
                    if self.req.get_current_program_scene() != "Game Capture":
                        self.req.set_current_program_scene("Game Capture")
                except:
                    pass
            
            time.sleep(interval)
            continue


    def get_emote(self):
        previous_emote = None
        while True:
            emote = choice(self.emoticons)
            if emote != previous_emote:
                yield emote
                previous_emote = emote
    
    def change_tabbed_text(self):
        interval = self._change_tabbed_text_poll_interval

        while True:
            self.check_connection()
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
        self.check_connection()

        t = Timer(2, self.change_tabbed_text)
        t.start()

        if self._change_tabbed_text_thread is None:
            self._change_tabbed_text_thread = threading.Thread(target=self.change_tabbed_text, daemon=True)
            self._change_tabbed_text_thread.start()
        
        """
        # start the notify.json monitor in background
        if self._notify_thread is None:
            self._notify_thread = threading.Thread(target=self.check_requested_add, daemon=True)
            self._notify_thread.start()
        """
        
        # start the switcher
        if self._switcher_thread is None:
            self._switcher_thread = threading.Thread(target=self.switcher, daemon=True)
            self._switcher_thread.start()

        # non blocking; keep main thread alive so the other children threads can do their jobs
        keepalive = Event()
        keepalive.wait()

if __name__ == "__main__":
    ms = MegaScript()
    ms.run()