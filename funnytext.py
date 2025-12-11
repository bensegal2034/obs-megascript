import obsws_python as obs
from threading import Timer
import asyncio

class FunnyText:

    def __init__(self):
        self.evt = obs.EventClient(host='localhost', port=4455, password='arcane2455', timeout=3)
        self.req = obs.ReqClient(host='localhost', port=4455, password='arcane2455', timeout=3)
        self.currentScene = None
        self.delay = 2

    def on_current_program_scene_changed(self, data):
        self.currentScene = data.scene_name
        print(self.currentScene)

    def change_text(self):
        print("blargh")
        if self.currentScene == "Alt Tabbed":
            print("hello!")

    def main(self):
        self.evt.callback.register(self.on_current_program_scene_changed)

        currentScene = self.req.get_current_program_scene().scene_name
        print(currentScene)

        t = Timer(self.delay, self.change_text())
        t.start()

        loop = asyncio.get_event_loop()
        loop.run_forever()

if __name__ == "__main__":
    ft = FunnyText()
    ft.main()