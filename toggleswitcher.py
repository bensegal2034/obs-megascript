import json

if __name__ == "__main__":
    commands_data = None
    with open("commands.json", "r") as f:
        commands_data = json.load(f)
    commands_data["toggleSwitcher"] = True
    with open("commands.json", "w") as f:
        json.dump(commands_data, f)