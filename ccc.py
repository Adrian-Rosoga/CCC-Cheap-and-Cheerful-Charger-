#!/usr/bin/env python3

"""

v1 - 26 Jan 2018
    # Support for TP-Link HS100 charger
v2 - 31 Dec 2019
    # Refactoring, encapsulate logic in Switch specific class

TODO
    # Ensure only one instance could run
    # requests instead of urllib
"""

import time
import os
import subprocess
import sys
import threading
import urllib.request
from urllib.error import URLError, HTTPError
import atexit
import signal
import logging
import argparse
import traceback
import platform
import psutil
from pathlib import Path
from playsound import playsound
import socket
from datetime import datetime
import datetime
from switch_plugins.switch import Switch
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

SLEEP_AFTER_INACTIVITY_MINS = 4

MIN_CHARGE, MAX_CHARGE = 45, 55
MIN_CHARGE_MANUAL, MAX_CHARGE_MANUAL = MIN_CHARGE - 1, MAX_CHARGE + 1
MAX_ALERT_CHARGE = 80
MIN_ALERT_CHARGE = 20
min_level = None
max_level = None

TIMEOUT = 10
START_QUIET_TIME = datetime.time(20, 0)
END_QUIET_TIME = datetime.time(6, 45)

switch = None
beep_only = False


def wifi_ssid() -> str:
    """Wifi ssid or empty string if not connected to wifi"""

    if IS_WINDOWS:
        cmd_list = ['netsh', 'wlan', 'show', 'interfaces']
    else:
        cmd_list = ['iwgetid', '-r']

    return subprocess.check_output(cmd_list).decode()


def should_be_quiet() -> bool:
    """At work keep it quiet"""
    return 'Barclays' in wifi_ssid()


def is_time_between(begin_time, end_time, check_time=None):
    check_time = check_time or datetime.datetime.now().time()
    if begin_time < end_time:
        return begin_time <= check_time <= end_time
    else:
        return check_time >= begin_time or check_time <= end_time


def voice_alert(soundfile):
    if not beep_only and not is_time_between(START_QUIET_TIME, END_QUIET_TIME):
        playsound(soundfile)


def beep(frequency=2500, duration_msec=1000):

    if should_be_quiet():
        return

    if IS_WINDOWS:
        winsound.Beep(frequency, duration_msec)
    else:
        playsound('beep-low-freq.wav')


def beep_loud(frequency=2500, duration_msec=3000):

    if should_be_quiet():
        return

    if IS_WINDOWS:
        winsound.Beep(frequency, duration_msec)
    else:
        playsound('beep-loud.wav')


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
            voice_alert('Battery_Low_Alert.wav')
        elif battery_level >= MAX_CHARGE_MANUAL and power_plugged():
            logging.info('Beep on battery_level > MAX_CHARGE_MANUAL and power_plugged()')
            beep(2000, 3000)
            voice_alert('Battery_High_Alert.wav')

    logging.info(f'{battery_level:.1f}% {switch.__class__.__name__} State={str(switch.state.name)} Charging={bool2onoff(power_plugged())}')

    if not control:
        return

    if battery_level <= min_level:

        if switch.state == Switch.State.OFF or switch.state == Switch.State.NA:
            switch.turn_on()
            logging.info(f'\t{battery_level:.1f}% - Turned power ON')

            # Check power is indeed on - wait for a few secs to give the power the time to reach the computer
            time.sleep(10)

        if not power_plugged():
            logging.error('\t### Not charging')
            beep(1000, 1000)
            voice_alert('Battery_Low_Alert.wav')

        # Turn power ON anyway to guard if the above command failed
        switch.turn_on()

        if switch.state == Switch.State.ON and not power_plugged():
            logging.warning('\t### Switch is ON but still not charging!')

    elif battery_level >= max_level:

        if switch.state != Switch.State.OFF or switch.state == Switch.State.NA:
            switch.turn_off()
            logging.info(f'\t{battery_level:.1f}% - Turned power OFF')

            # Check power is indeed off - wait a few secs
            time.sleep(10)
            if power_plugged():
                logging.error('\t### Switch stuck on ON position!?')
                beep(1000, 3000)
                voice_alert('Battery_High_Alert.wav')

        # Turn power OFF anyway to guard if the above command failed
        switch.turn_off()


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
        logging.info(f'hwnd={hwnd}')
    except Exception as e:
        logging.error(f'Exception caught: {str(e)}')
        return

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
            except socket.timeout:
                logging.error('socket.timeout: Socket timed out')
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
                logging.info(f'\t### Overcharged above {MAX_ALERT_CHARGE:.1f}% - at {battery_level:.1f}%')
                if power_plugged():
                    beep_loud(500, 3000)
                    time.sleep(0.1)
                    beep_loud(500, 3000)
                    time.sleep(0.1)
                    beep_loud(500, 3000)
            elif battery_level <= MIN_ALERT_CHARGE:
                logging.info(f'\t### Undercharged below {MIN_ALERT_CHARGE:.1f}% - at {battery_level:.1f}%')
                if not power_plugged():
                    beep_loud(500, 3000)

            time.sleep(60)


class SingleInstanceThread(threading.Thread):

    def run(self):

        if IS_WINDOWS:

            logging.info("Checking if another instance is running...")

            mutex = win32event.CreateMutex(None, False, 'ccc')
            last_error = win32api.GetLastError()

            if last_error == ERROR_ALREADY_EXISTS:
                logging.info('Another instance already running. Exiting.')
                beep_loud()
                os._exit(1)

            while True:
                time.sleep(10000)


class SleepThread(threading.Thread):

    def run(self):

        while True:

            output = subprocess.check_output(['./xprintidle'])
            inactivity_secs = int(output.decode()) // 1000

            if inactivity_secs >= SLEEP_AFTER_INACTIVITY_MINS * 60:
                logging.info(f'No user activity in the last {inactivity_secs} seconds. Turning power off and going to sleep...')

                try:
                    switch.turn_off()
                except:
                    logging.error('Exception thrown when turning the switch off')
                
                time.sleep(5)
                
                # Suspend only if power is indeed disconnected, else keep it awake to at least alert in case
                # it would overcharge
                if not power_plugged():
                    os.system('systemctl suspend')

            else:
                logging.info(f'User activity detected {inactivity_secs} secs ago. Staying awake!')

            # Sleep a few secs longer than the inactivity interval
            time.sleep(SLEEP_AFTER_INACTIVITY_MINS * 60 + 10)


def test_on_off():

    while True:
        switch.turn_on()
        time.sleep(30)
        switch.turn_off()
        time.sleep(30)


def has_battery():
    return psutil.sensors_battery() is not None


def main():

    if not has_battery():
        print('No battery detected. This program won\'t be of any help. Exiting.')
        return 1
  
    log_file = Path.home() / LOG_FILE

    handlers = [logging.FileHandler(log_file), logging.StreamHandler()]

    logging.basicConfig(format="%(asctime)-15s - %(message)s",
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.INFO,
                        handlers=handlers)

    parser = argparse.ArgumentParser(description='CCC (Cheap and Cheerful Charger)')
    parser.add_argument("switch", help="type of switch: hs100, energenie, noswitch")
    parser.add_argument("--min", help="minimum charge for alert")
    parser.add_argument("--max", help="maximum charge for alert")
    parser.add_argument('--nocontrol', help='no power control, just monitor', action='store_true')
    parser.add_argument('--inactivity', help='make computer sleep on inactivity', action='store_true')
    parser.add_argument('--beep', help='beep only', action='store_true')
    args = parser.parse_args()

    global switch
    global beep_only
    global min_level, max_level
    global MAX_ALERT_CHARGE, MIN_ALERT_CHARGE

    if args.switch == 'noswitch':
        switch = NoSwitch()
    elif args.switch == 'energenie':
        switch = EnergenieSwitch(TIMEOUT)
    elif args.switch == 'hs100':
        switch = HS100Switch()
    else:
        print('Invalid switch value. Exiting.')
        return 1

    #test_on_off()

    sleep_on_inactivity = args.inactivity
    control = not args.nocontrol
    beep_only = args.beep

    print('')

    if not control:
        switch = NoSwitch()
        print('Monitoring mode, power source not controlled')
        min_level = MIN_CHARGE_MANUAL
        max_level = MAX_CHARGE_MANUAL
    else:
        min_level = MIN_CHARGE
        max_level = MAX_CHARGE

    if args.min is not None:
        min_level = int(args.min)
        MIN_ALERT_CHARGE = min_level
    if args.max is not None:
        max_level = int(args.max)
        MAX_ALERT_CHARGE = max_level

    print('************* Cheap and Cheerful Charger *************')
    print(f'Arguments: {" ".join(sys.argv[1:])}')
    print(f'Charge range: ({min_level}% - {max_level}%)')
    print(f'Logging to: {log_file}')

    print('******************************************************')
    print('')

    sys.stderr = sys.stdout

    # To prevent overcharging turn power off when killing the program
    if not IS_WINDOWS:
        atexit.register(turn_power_off)
        signal.signal(signal.SIGUSR1, handler)

    SingleInstanceThread().start()

    PowerControlThread(control).start()

    WatchdogThread().start()

    if sleep_on_inactivity:
        if not IS_WINDOWS:
            SleepThread().start()

    if False and IS_WINDOWS:
        listen_for_sleep()


if __name__ == "__main__":
    main()
