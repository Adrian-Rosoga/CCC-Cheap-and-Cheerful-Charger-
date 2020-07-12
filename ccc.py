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
import psutil
from playsound import playsound
import socket
from struct import pack
import parser
from switch_plugins.no_switch import NoSwitch
from switch_plugins.hs100_switch import HS100Switch
from switch_plugins.energenie_switch import EnergenieSwitch


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

MIN_CHARGE, MAX_CHARGE = 45, 55
MIN_CHARGE_MANUAL, MAX_CHARGE_MANUAL = MIN_CHARGE - 1, MAX_CHARGE + 1
MAX_ALERT_CHARGE = MAX_CHARGE + 5
MIN_ALERT_CHARGE = MIN_CHARGE - 5

IP_PORT = '192.168.1.157:9999'

switch = None
beep_only = False


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


def voice_alert(soundfile):
    if not beep_only:
        playsound(soundfile)


def beep(frequency=2500, duration_msec=1000):

    if should_be_quiet():
        return

    if IS_WINDOWS:
        winsound.Beep(frequency, duration_msec)
    else:
        playsound('beep-low-freq.wav')


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
            voice_alert('Battery_Low_Alert.wav')
        elif battery_level >= MAX_CHARGE_MANUAL and power_plugged():
            logging.info('Beep on battery_level > MAX_CHARGE_MANUAL and power_plugged()')
            beep(2000, 3000)
            voice_alert('Battery_High_Alert.wav')

    logging.info(f'{battery_level:.1f}% {switch.__class__.__name__} State={str(switch.state.name)} Charging={bool2onoff(power_plugged())}')

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
            voice_alert('Battery_Low_Alert.wav')

        # Turn power ON anyway to guard if the above command failed
        switch.turn_on()

        if switch.state == Switch.State.ON and not power_plugged():
            logging.warning('\t### Switch is ON but still not charging!')

    elif battery_level >= MAX_CHARGE:

        if switch.state != Switch.State.OFF or switch.state == Switch.State.NA:
            switch.turn_off()
            logging.info(f'\t{battery_level:.1f}% - Turned power OFF')

            # Check power is indeed off - wait a few secs
            time.sleep(10)
            if power_plugged():
                logging.error('\t### Switch stuck on ON position!?')
                beep(1000, 1000)
                voice_alert('Battery_High_Alert.wav')

        # Turn power OFF anyway to guard if the above command failed
        switch.turn_off()

        if power_plugged():
            logging.warning('\t### Charging above the threshold!')
            beep(1000, 1000)
            voice_alert('Battery_High_Alert.wav')


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

        SLEEP_AFTER_MINS = 2
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


def has_battery():
    return psutil.sensors_battery() is not None


def main():

    print('\n=== Cheap and Cheerful Charger ===\n')

    if not has_battery():
        print('No battery detected. This program won\'t be of any help. Exiting.')
        return 1

    logging.basicConfig(format="%(asctime)-15s - %(message)s",
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.INFO)

    parser = argparse.ArgumentParser(description='CCC (Cheap and Cheerful Charger)')
    parser.add_argument('--nocontrol', help='no power control, just monitor', action='store_true')
    parser.add_argument('--inactivity', help='make computer sleep on inactivity', action='store_true')
    parser.add_argument('--beep', help='beep only', action='store_true')
    args = parser.parse_args()

    global switch
    global beep_only

    switch = NoSwitch()
    #switch = EnergenieSwitch()
    #switch = HS100Switch()

    #test_on_off()

    sleep_on_inactivity = args.inactivity
    control = not args.nocontrol
    beep_only = args.beep

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
