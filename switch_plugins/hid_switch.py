from switch_plugins.switch import Switch


class HIDSwitch(Switch):

    @property
    def state(self):
        output = subprocess.check_output(['hidusb-Switch-cmd.exe', 'state'])
        if 'R1=ON' in output.decode():
            return Switch.State.ON
        elif 'R1=OFF' in output.decode():
            return Switch.State.OFF
        else:
            return Switch.State.NA

    def turn_on(self):
        os.system('hidusb-relay-cmd.exe on 1')

    def turn_off(self):
        os.system('hidusb-relay-cmd.exe off 1')
