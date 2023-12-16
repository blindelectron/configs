import addonHandler
import config
from enum import Enum
import re
addonHandler.initTranslation()


class PatternType(Enum):
    MESSAGE = "^{} (?P<userName>.*?)[:.] (?P<text>.*)$"

events = [
    {"name": "chanmsg", "label": _("Channel messages"), "text": "Channel message from", "pattern": PatternType.MESSAGE.value},
    {"name": "usermsg", "label": _("User messages"), "text": "(Text|Private) message from", "pattern": PatternType.MESSAGE.value},
    {"name": "bcastmsg", "label": _("Broadcast messages"), "text": "Broadcast message from", "pattern": PatternType.MESSAGE.value}
]

def getAllPatterns():
    addonInstance = addonHandler.getCodeAddon()
    addonName = addonInstance.name
    addonConf = config.conf[addonName]
    patterns = []
    for event in events:
        text = _(event["text"]) if addonConf["regexTranslation"] else event["text"]
        pattern = {"name": event["name"], "re": re.compile(event["pattern"].format(text), flags=re.DOTALL)}
        patterns.append(pattern)
    return patterns
