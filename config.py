import json
CONFIG = {}
DN_RECORD = []
with open("config.json", "r") as config:
    CONFIG = json.load(config)


def update_config() -> None:
    with open("config.json", "w") as config:
        json.dump(CONFIG, config)
