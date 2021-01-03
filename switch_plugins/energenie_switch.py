import urllib
from switch_plugins.switch import Switch

class EnergenieSwitch(Switch):

    def __init__(self, timeout):
        self.timeout = timeout

    @property
    def state(self):
        return Switch.State.NA

    def turn_on(self):
        urllib.request.urlopen("http://192.168.1.108:8000/on", timeout=self.timeout)

    def turn_off(self):
        urllib.request.urlopen("http://192.168.1.108:8000/off", timeout=self.timeout)
