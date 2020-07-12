from switch_plugins.switch import Switch

class EnergenieSwitch(Switch):

    @property
    def state(self):
        return Switch.State.NA

    def turn_on(self):
        urllib.request.urlopen("http://192.168.1.108:8000/on", timeout=TIMEOUT)

    def turn_off(self):
        urllib.request.urlopen("http://192.168.1.108:8000/off", timeout=TIMEOUT)