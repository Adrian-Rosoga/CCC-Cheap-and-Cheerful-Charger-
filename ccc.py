#!/usr/bin/env python3

'''

v1 - 26 Jan 2018
    # Support for TP-Link HS100 charger
v2 - 31 Dec 2019
    # Refactoring, encapsulate logic in Switch specific class

TODO
    # Ensure only one instance could run
    # requests instead of urllib
'''


import time
import os
import subprocess
import sys
import threading
import urllib.request
import json
import atexit
import signal
import logging
import argparse
import traceback
import platform
from socket import timeout
from urllib.error import URLError, HTTPError
from abc import ABC, abstractmethod
import psutil
from playsound import playsound
from enum import Enum, unique, auto
import socket
from struct import pack
import parser


IS_WINDOWS = platform.system() == 'Windows'


if IS_WINDOWS:
    import winsound
    import win32con
    import win32api
    import win32gui
    import win32event
    from winerror import ERROR_ALREADY_EXISTS


LOG_FILE = 'ccc.log'
TIMEOUT = 10

MIN_CHARGE, MAX_CHARGE = 25, 75
MIN_CHARGE, MAX_CHARGE = 45, 55

MIN_CHARGE_MANUAL, MAX_CHARGE_MANUAL = MIN_CHARGE - 1, MAX_CHARGE + 1

MAX_ALERT_CHARGE = MAX_CHARGE + 5
MIN_ALERT_CHARGE = MIN_CHARGE - 5

# Defining static vars
IP = "192.168.1.157" #  Checks IP is valid, change to your smart-plug IP
PORT = 9999 #  9999 is default port. Change if need to

switch = None


# Encrypts value to be sent
def encrypt(string):
    key = 171
    result = pack('>I', len(string))
    for i in string:
        a = key ^ ord(i)
        key = a
        result += bytes([a])
    return result


# Decrypts return value
def decrypt(string):
    key = 171
    result = ""
    for i in string:
        a = key ^ i
        key = i
        result += chr(a)
    return result


# Basic commands
commands = {'info'     : '{"system":{"get_sysinfo":{}}}',
            'on'       : '{"system":{"set_relay_state":{"state":1}}}',
            'off'      : '{"system":{"set_relay_state":{"state":0}}}',
            'ledoff'   : '{"system":{"set_led_off":{"off":1}}}',
            'ledon'    : '{"system":{"set_led_off":{"off":0}}}',
            'cloudinfo': '{"cnCloud":{"get_info":{}}}',
            'wlanscan' : '{"netif":{"get_scaninfo":{"refresh":0}}}',
            'time'     : '{"time":{"get_time":{}}}',
            'schedule' : '{"schedule":{"get_rules":{}}}',
            'countdown': '{"count_down":{"get_rules":{}}}',
            'antitheft': '{"anti_theft":{"get_rules":{}}}',
            'reboot'   : '{"system":{"reboot":{"delay":1}}}',
            'reset'    : '{"system":{"reset":{"delay":1}}}',
            'energy'   : '{"emeter":{"get_realtime":{}}}'
}

# Sends command to device
def sendCommand(cmd):
    try:
        sock_tcp = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock_tcp.settimeout(TIMEOUT)
        sock_tcp.connect((IP, PORT))
        sock_tcp.settimeout(None)
        sock_tcp.send(encrypt(cmd))
        data = sock_tcp.recv(2048)
        sock_tcp.close()
        return data
    except socket.error:
        logging.error("Could not connect to host " + IP + ":" + str(PORT))


def wifi_ssid() -> str:
    """Wifi ssid or empty string if not connected to wifi"""

    if IS_WINDOWS:
        cmd_list = ['netsh', 'wlan', 'show', 'interfaces']
    else:
        cmd_list = ['iwgetid', '-r']

    output = subprocess.check_output(cmd_list)

    return output.decode()


def should_be_quiet() -> bool:
    """At work keep it quiet"""

    return 'Barclays' in wifi_ssid()


def beep(frequency=2500, duration_msec=1000):

    if should_be_quiet():
        return

    if IS_WINDOWS:
        winsound.Beep(frequency, duration_msec)
    else:
        playsound('beep-low-freq.wav')


class Switch(ABC):

    @unique
    class State(Enum):
        ON = auto()
        OFF = auto()
        NA = auto()

    @property
    @abstractmethod
    def state(self):
        pass

    @abstractmethod
    def turn_on(self):
        pass

    @abstractmethod
    def turn_off(self):
        pass


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


class HS100Switch(Switch):

    @property
    def state(self):
        response = sendCommand(commands["info"])
        if response is None:    # Because on holiday for example
            return Switch.State.NA
        info = decrypt(response)
        info = '{' + info[5:]
        data = json.loads(info)
        return Switch.State.ON if data['system']['get_sysinfo']['relay_state'] == 1 else Switch.State.OFF

    def turn_on(self):
        sendCommand(commands["on"])

    def turn_off(self):
        sendCommand(commands["off"])


class EnergenieSwitch(Switch):

    @property
    def state(self):
        return Switch.State.NA

    def turn_on(self):
        urllib.request.urlopen("http://192.168.1.108:8000/on", timeout=TIMEOUT)

    def turn_off(self):
        urllib.request.urlopen("http://192.168.1.108:8000/off", timeout=TIMEOUT)


class NoSwitch(Switch):

    @property
    def state(self):
        return Switch.State.NA

    def turn_on(self):
        pass

    def turn_off(self):
        pass


def power_plugged():
    return psutil.sensors_battery().power_plugged


def turn_power_off():
    logging.info('Program terminating, turning power off...')
    switch.turn_off()


def handler(signum, frame):
    logging.info(f'Signal handler called with signal {signal.Signals(signum).name}')
    if signum == signal.SIGUSR1:
        turn_power_off()


def battery_percent():
    return psutil.sensors_battery().percent


def bool2onoff(value):
    return 'ON' if value else 'OFF'


def control(control=True):

    global switch

    battery_level = battery_percent()

    if False:
        # Hack for manual charging
        if battery_level <= MIN_CHARGE_MANUAL and not power_plugged():
            logging.info('Beep on battery_level < MIN_CHARGE_MANUAL and not power_plugged()')
            beep(1000, 1000)
            playsound('Battery_Low_Alert.wav')
        elif battery_level >= MAX_CHARGE_MANUAL and power_plugged():
            logging.info('Beep on battery_level > MAX_CHARGE_MANUAL and power_plugged()')
            beep(2000, 3000)
            playsound('Battery_High_Alert.wav')

    logging.info(f'{battery_level:.1f}% {switch.__class__.__name__} State={str(switch.state.name)} Power={bool2onoff(power_plugged())}')

    if not control:
        return

    if battery_level <= MIN_CHARGE:

        if switch.state == Switch.State.OFF or switch.state == Switch.State.NA:
            switch.turn_on()
            logging.info(f'\t{battery_level:.1f}% - Turned power ON')

            # Check power is indeed on - wait for a few secs to give the power the time to reach the computer
            time.sleep(10)

        if not power_plugged():
            logging.error('\t### Not charging')
            beep(1000, 1000)  # Get rid of the pesky 2 beeps
            playsound('Battery_Low_Alert.wav')

        # Turn power ON anyway to guard if the above command failed
        switch.turn_on()

        if switch.state == Switch.State.ON and not power_plugged():
            logging.warning('\t### Switch is ON but still not charging!')
            beep(1000, 1000)

    elif battery_level >= MAX_CHARGE:

        if switch.state != Switch.State.OFF or switch.state == Switch.State.NA:
            switch.turn_off()
            logging.info(f'\t{battery_level:.1f}% - Turned power OFF')

            # Check power is indeed off - wait a few secs
            time.sleep(10)
            if power_plugged():
                logging.error('\t### Switch stuck on ON position!?')
                beep(1000, 1000)
                playsound('Battery_High_Alert.wav')

        # Turn power OFF anyway to guard if the above command failed
        switch.turn_off()

        if switch.state == Switch.State.OFF and power_plugged():
            logging.warning('\t### Switch is OFF but still charging')
            beep(1000, 1000)


def wndproc(hwnd, msg, wparam, lparam):
    '''
    At logoff:
    wndproc: msg=17 wparam=0 lparam=-2147483648
    wndproc: msg=22 wparam=1 lparam=-2147483648
    '''

    #logging.info("wndproc: msg=%s wparam=%s lparam=%s" % (msg, wparam, lparam))
    if (msg == win32con.WM_POWERBROADCAST and wparam == win32con.PBT_APMSUSPEND) or \
       (msg == win32con.WM_ENDSESSION):

        switch.turn_off()
        logging.info(f'\t### {battery_percent():1.f}% - Entering sleep, turned power OFF')


def listen_for_sleep():

    hinst = win32api.GetModuleHandle(None)
    wndclass = win32gui.WNDCLASS()
    wndclass.hInstance = hinst
    wndclass.lpszClassName = "dummy_window"
    messageMap = {win32con.WM_QUERYENDSESSION: wndproc,
                  win32con.WM_ENDSESSION: wndproc,
                  win32con.WM_QUIT: wndproc,
                  win32con.WM_DESTROY: wndproc,
                  win32con.WM_CLOSE: wndproc,
                  win32con.WM_POWERBROADCAST: wndproc}

    wndclass.lpfnWndProc = messageMap

    try:
        myWindowClass = win32gui.RegisterClass(wndclass)
        hwnd = win32gui.CreateWindowEx(win32con.WS_EX_LEFT,
                                       myWindowClass,
                                       "dummy_window",
                                       0,
                                       0,
                                       0,
                                       win32con.CW_USEDEFAULT,
                                       win32con.CW_USEDEFAULT,
                                       0,
                                       0,
                                       hinst,
                                       None)
    except Exception as e:
        logging.error(f'Exception caught: {str(e)}')

    logging.info(f'hwnd={hwnd}')

    while True:
        win32gui.PumpWaitingMessages()
        time.sleep(1)


class PowerControlThread(threading.Thread):

    def __init__(self, control):
        super().__init__()
        self.control = control

    def run(self):

        while True:

            try:
                control(self.control)
            except HTTPError as e:
                logging.error('HTTPError: The server couldn\'t fulfill the request')
                logging.error('\tError code: ' + str(e.code))
            except URLError as e:
                logging.error('URLError: We failed to reach a server')
                logging.error('\tReason: ' + str(e.reason))
            except timeout:
                logging.error('timeout: socket timed out')
            except Exception as ex:
                logging.error('Exception: ' + ex.__class__.__name__)
                logging.error(ex)
                traceback.print_exc(file=sys.stdout)

            time.sleep(60)


class WatchdogThread(threading.Thread):

    def run(self):

        while True:

            battery_level = battery_percent()
            if battery_level >= MAX_ALERT_CHARGE:
                logging.info(f'\t### Overcharged above {MAX_ALERT_CHARGE:.1f}% - {battery_level:.1f}%')
                if power_plugged():
                    beep(500, 3000)
                    time.sleep(0.1)
                    beep(500, 3000)
                    time.sleep(0.1)
                    beep(500, 3000)
            elif battery_level <= MIN_ALERT_CHARGE:
                logging.info(f'\t### Undercharged below {MIN_ALERT_CHARGE:.1f}% - {battery_level:.1f}%')
                if not power_plugged():
                    beep(500, 3000)

            time.sleep(60)


class SingleInstanceThread(threading.Thread):

    def run(self):

        if IS_WINDOWS:

            logging.info("Checking if another instance is running...")

            mutex = win32event.CreateMutex(None, False, 'ccc')
            last_error = win32api.GetLastError()

            if last_error == ERROR_ALREADY_EXISTS:
                logging.info('Another instance already running. Exiting.')
                os._exit(1)

            while True:
                time.sleep(10000)


class SleepThread(threading.Thread):

    def run(self):

        SLEEP_AFTER_MINS = 5
        SLEEP_AFTER_SECS = SLEEP_AFTER_MINS * 60

        while True:

            output = subprocess.check_output(['xprintidle'])
            inactivity_secs = int(output.decode()) // 1000

            if inactivity_secs >= SLEEP_AFTER_SECS:
                logging.info(f'No user activity in the last {inactivity_secs} seconds. Turning power off and going to sleep...')

                try:
                    switch.turn_off()
                except:
                    logging.error('Exception thrown when turning the switch off')
                
                time.sleep(5)
                os.system('systemctl suspend')

            else:
                logging.info(f'User activity detected {inactivity_secs} secs ago. Staying awake!')

            time.sleep(SLEEP_AFTER_SECS + 10)


def test_on_off():

    while True:
        switch.turn_on()
        time.sleep(30)
        switch.turn_off()
        time.sleep(30)


def main():

    print('\n=== Cheap and Cheerful Charger ===\n')

    logging.basicConfig(format="%(asctime)-15s - %(message)s",
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.INFO)

    parser = argparse.ArgumentParser(description='CCC (Cheap and Cheerful Charger)')
    parser.add_argument('--nocontrol', help='no power control, just monitor', action='store_true')
    parser.add_argument('--inactivity', help='make computer sleep on inactivity', action='store_true')
    args = parser.parse_args()

    global switch

    #switch = EnergenieSwitch()
    switch = HS100Switch()

    #test_on_off()

    sleep_on_inactivity = args.inactivity

    control = not args.nocontrol

    logging.info('*****************************************************')
    logging.info('*****************************************************')

    if not control:
        switch = NoSwitch()
        logging.info('Monitoring mode, power source not controlled')
        min_level = MIN_CHARGE_MANUAL
        max_level = MAX_CHARGE_MANUAL
    else:
        min_level = MIN_CHARGE
        max_level = MAX_CHARGE

    logging.info(f'Charge range is ({min_level}% - {max_level}%)')

    logging.info('*****************************************************')
    logging.info('*****************************************************')

    sys.stderr = sys.stdout

    # To prevent overcharging tunr power off when killing the program
    if not IS_WINDOWS:
        atexit.register(turn_power_off)
        signal.signal(signal.SIGUSR1, handler)

    SingleInstanceThread().start()

    PowerControlThread(control).start()

    if False:
        WatchdogThread().start()

    if sleep_on_inactivity:
        if not IS_WINDOWS:
            SleepThread().start()

    if False and IS_WINDOWS:
        listen_for_sleep()


if __name__ == "__main__":
    main()
