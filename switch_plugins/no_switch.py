from switch_plugins.switch import Switch

class NoSwitch(Switch):

    @property
    def state(self):
        return Switch.State.NA

    def turn_on(self):
        pass

    def turn_off(self):
        pass