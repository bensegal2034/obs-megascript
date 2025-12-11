import obsws_python as obs
from threading import Timer, Event
from random import choice
from playsound import playsound
from ctypes import wintypes, windll, create_unicode_buffer

# PLAN TO IMPLEMENT AUTO SCENE CHANGER
# have a json file which is a list of all of the windows we care about with an attribute
# that dictates which scene to switch to when the window is active
# when one of those windows is the active window, change the scene accordingly

# in order to add to this json file, i will create a second script that will be ran when a stream deck button is hit
# this script will update a txt file to notify this script the button has been hit
# this script then clears the txt file and adds the currently active window to the list

class MegaScript:

    def __init__(self):
        self.evt = obs.EventClient(host='localhost', port=4455, password='arcane2455', timeout=3)
        self.req = obs.ReqClient(host='localhost', port=4455, password='arcane2455', timeout=3)
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

    def on_replay_buffer_saved(self, data):
        print("Detected replay buffer save")
        playsound("D:\\Music\\recordingbeep.mp3")
    
    def getForegroundWindowTitle():
        hWnd = windll.user32.GetForegroundWindow()
        length = windll.user32.GetWindowTextLengthW(hWnd)
        buf = create_unicode_buffer(length + 1)
        windll.user32.GetWindowTextW(hWnd, buf, length + 1)
        
        # 1-liner alternative: return buf.value if buf.value else None
        if buf.value:
            return buf.value
        else:
            return None

    def get_emote(self):
        previous_emote = None
        while True:
            emote = choice(self.emoticons)
            if emote != previous_emote:
                yield emote
                previous_emote = emote
    
    def change_tabbed_text(self):
        if self.req.get_current_program_scene().scene_name == "Alt Tabbed":
            self.req.set_input_settings("Alt Tabbed Text", {
                "text": f"Alt Tabbed {next(self.emote_gen)}"
            }, True)
        t = Timer(2, self.change_tabbed_text)
        t.start()

    def run(self):
        self.evt.callback.register(self.on_replay_buffer_saved)

        t = Timer(2, self.change_tabbed_text)
        t.start()

        # non blocking
        keepalive = Event()
        keepalive.wait()

if __name__ == "__main__":
    ms = MegaScript()
    ms.run()