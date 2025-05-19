import json
import re
CONFIG = {}
DN_RECORD = []

SERIAL_REGEX = "158\d{10}"

with open("config.json", "r") as config:
    CONFIG = json.load(config)


def update_config() -> None:
    with open("config.json", "w") as config:
        json.dump(CONFIG, config)
