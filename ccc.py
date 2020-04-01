#!/usr/bin/env python3

'''

v1 - 26 Jan 2018
    # Support for TP-Link HS100 charger
v2 - 31 Dec 2019
    # Refactoring, encapsulate logic in relay specific class

TODO
    # Ensure only one instance could run
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
from contextlib import AbstractContextManager
import psutil
from pygame import mixer


IS_WINDOWS = platform.system() == 'Windows'


if IS_WINDOWS:
    import win32con
    import win32api
    import win32gui
    import win32event
    from winerror import ERROR_ALREADY_EXISTS


#LOG_FILE = 'C:\\Tmp\\ccc.log'
LOG_FILE = 'ccc.log'

MIN_CHARGE, MAX_CHARGE = 35, 65
#MIN_CHARGE, MAX_CHARGE = 45, 55
#MIN_CHARGE, MAX_CHARGE = 49, 51
#MIN_CHARGE, MAX_CHARGE = 58, 60

#MIN_CHARGE_MANUAL, MAX_CHARGE_MANUAL = 40, 60
MIN_CHARGE_MANUAL, MAX_CHARGE_MANUAL = MIN_CHARGE - 1, MAX_CHARGE + 1
MAX_ALERT_CHARGE = MAX_CHARGE + 5
MIN_ALERT_CHARGE = MIN_CHARGE - 5

ON = True
OFF = False
TIMEOUT = 10


class Beeper(AbstractContextManager):

    def __enter__(self):
        mixer.init()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        mixer.quit()

    @staticmethod
    def beep(soundfile='beep-low-freq.wav', duration_secs=1):
        mixer.Sound(soundfile).play()
        time.sleep(duration_secs)


def wifi_ssid() -> str:
    """Return the wifi ssid or empty string if not connected to wifi"""

    if not IS_WINDOWS:
        popen = subprocess.Popen(['iwgetid'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, _ = popen.communicate()
    else:
        popen = subprocess.Popen(['netsh', 'wlan', 'show', 'interfaces'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, _ = popen.communicate()
        
    print(output.decode())
    return output.decode()  


def should_be_quiet() -> bool:
    """At work keep it quiet"""
    
    #return 'Barclays' in wifi_ssid()
    wifi_ssid()
    return False


def beep(frequency=2500, duration_msec=1000):

    if should_be_quiet():
        return

    if IS_WINDOWS:
        import winsound
        winsound.Beep(frequency, duration_msec)
    else:
        with Beeper() as beeper:
            beeper.beep()


class Relay(ABC):

    @property
    @abstractmethod
    def state(self) -> str:
        return ""

    @abstractmethod
    def turn_power(self, on_off: bool) -> None:
        pass


class HIDRelay(Relay):

    @property
    def state(self) -> str:

        popen = subprocess.Popen(['hidusb-relay-cmd.exe', 'state'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, _ = popen.communicate()
        if 'R1=ON' in output.decode("utf-8"):
            return 'ON'
        elif 'R1=OFF' in output.decode("utf-8"):
            return 'OFF'
        else:
            return 'N/A'

    def turn_power(self, on_off: bool) -> None:

        if on_off == ON:
            os.system('hidusb-relay-cmd.exe on 1')
        else:
            os.system('hidusb-relay-cmd.exe off 1')


class HS100Relay(Relay):

    @property
    def state(self) -> str:

        output = urllib.request.urlopen("http://192.168.1.103/status", timeout=TIMEOUT).read()
        data = json.loads(output)
        #print(output)
        if data['system']['get_sysinfo']['state'] is None:
            return 'N/A'
        return 'ON' if data['system']['get_sysinfo']['state'] == 1 else 'OFF'

    def turn_power(self, on_off: bool) -> None:

        if on_off == ON:
            urllib.request.urlopen("http://192.168.1.103/on", timeout=TIMEOUT)
        else:
            urllib.request.urlopen("http://192.168.1.103/off", timeout=TIMEOUT)


class EnergenieRelay(Relay):

    @property
    def state(self) -> str:
        return 'N/A'

    def turn_power(self, on_off: bool) -> None:

        if on_off == ON:
            urllib.request.urlopen("http://192.168.1.108:8000/on", timeout=TIMEOUT)
        else:
            urllib.request.urlopen("http://192.168.1.108:8000/off", timeout=TIMEOUT)


relay = EnergenieRelay()


def power_plugged():

    value = psutil.sensors_battery().power_plugged
    if value:
        return 'ON'
    elif not value:
        return 'OFF'
    else:
        logging.warning('\t### Cannot determine if plugged in or not')
        return 'OFF'


def turn_power_off():

    try:
        logging.info('Program getting killed, turning power off now...')
        relay.turn_power(OFF)
        logging.info('Program killed, turned power off')
    except urllib.error.URLError:
        logging.error('Caught urllib.error.URLError')
    except OSError:
        logging.error('Caught OSError')
    except Exception:
        print('TODO: Exception caught in turn_power_off()')
        logging.error("Something bad happened", exc_info=True)


def handler(signum, frame):
    logging.info(f'Signal handler called with signal {signal.Signals(signum).name}')
    if signum == signal.SIGUSR1:
        turn_power_off()


def battery_percent():
    return psutil.sensors_battery().percent


def control():

    # Hack for manual charging
    if battery_percent() < MIN_CHARGE_MANUAL and power_plugged() == 'OFF':
        beep(1000, 1000)
    elif battery_percent() > MAX_CHARGE_MANUAL and power_plugged() == 'ON':
        beep(2000, 3000)

    logging.info(f'{battery_percent():.1f}% Relay={relay.state} Power={power_plugged()}')

    if relay.state == 'N/A':
        #return
        pass

    if battery_percent() <= MIN_CHARGE:

        if relay.state == 'OFF' or relay.state == 'N/A':
            relay.turn_power(ON)
            logging.info(f'\t{battery_percent():.1f}% - Turn power ON')

            # Check power is indeed on - wait for a few secs to give the power the time to reach the
            # computer
            time.sleep(3)
            if power_plugged() != 'ON':
                logging.error('\t### Not charging although power turned ON')
                # beep(1000, 1000)  # Get rid of the pesky 2 beeps

        # Turn power ON anyway to guard if the above command failed
        relay.turn_power(ON)

    elif battery_percent() >= MAX_CHARGE:

        if relay.state != 'OFF' or relay.state == 'N/A':
            relay.turn_power(OFF)
            logging.info(f'\t{battery_percent():.1f}% - Turn power OFF')

            # Check power is indeed off - wait a few secs
            time.sleep(3)
            if power_plugged() == 'ON':
                logging.error('\t### Relay stuck on ON position!?')
                beep(1000, 1000)

        # Turn power OFF anyway to guard if the above command failed
        relay.turn_power(OFF)

    if relay.state == 'OFF' and power_plugged() == 'ON':
        logging.warning('\t### Charging when not supposed to!?')
        beep(1000, 1000)

    if relay.state == 'ON' and power_plugged() == 'OFF':
        logging.warning('\t### Plug charger in or check why not charging!')
        beep(1000, 1000)


def test_on_off():

    while True:
        relay.turn_power(ON)
        time.sleep(30)
        relay.turn_power(OFF)
        time.sleep(30)


'''
At logoff:
wndproc: msg=17 wparam=0 lparam=-2147483648
wndproc: msg=22 wparam=1 lparam=-2147483648
'''


def wndproc(hwnd, msg, wparam, lparam):

    #logging.info("wndproc: msg=%s wparam=%s lparam=%s" % (msg, wparam, lparam))
    if (msg == win32con.WM_POWERBROADCAST and wparam == win32con.PBT_APMSUSPEND) or \
       (msg == win32con.WM_ENDSESSION):

        relay.turn_power(OFF)
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
        logging.info("Exception: %s" % str(e))

    if hwnd is None:
        logging.info("hwnd is none!")
    else:
        logging.info("hwnd=%s" % hwnd)
        pass

    while True:
        win32gui.PumpWaitingMessages()
        time.sleep(1)


class PowerControlThread(threading.Thread):

    def __init__(self):

        threading.Thread.__init__(self)

    def run(self):

        while True:

            try:
                control()
            except HTTPError as e:
                logging.info('HTTPError: The server couldn\'t fulfill the request')
                logging.info('\tError code: ' + str(e.code))
                #beep(600, 2000)
            except URLError as e:
                logging.info('URLError: We failed to reach a server')
                logging.info('\tReason: ' + str(e.reason))
                #beep(600, 2000)
            except timeout:
                logging.info('timeout: socket timed out')
                #beep(600, 2000)
            except Exception as ex:
                logging.info('Exception: ' + ex.__class__.__name__)
                print(ex)
                traceback.print_exc(file=sys.stdout)
                #beep(600, 2000)

            time.sleep(60)


class WatchdogThread(threading.Thread):

    def __init__(self):

        threading.Thread.__init__(self)

    def run(self):

        while True:

            if battery_percent() >= MAX_ALERT_CHARGE:
                logging.info(f'\t### Overcharged above {MAX_ALERT_CHARGE:.1f}% - {battery_percent():1.f}%')
                if relay.state == 'ON':
                    beep(500, 3000)

            if battery_percent() <= MIN_ALERT_CHARGE:
                logging.info(f'\t### Undercharged below {MIN_ALERT_CHARGE:.1f}% - {battery_percent():.1f}%')
                
                if relay.state == 'OFF':
                    beep(500, 3000)

            time.sleep(60)


class SingleInstanceThread(threading.Thread):

    def __init__(self):

        threading.Thread.__init__(self)

    def run(self):

        if IS_WINDOWS:

            print("Checking if another instance is running...")

            mutex = win32event.CreateMutex(None, False, 'ccc')
            last_error = win32api.GetLastError()

            if last_error == ERROR_ALREADY_EXISTS:
                print('Another instance already running. Exiting.')
                os._exit(1)

            while True:
                time.sleep(10000)


def main():

    

    logging.basicConfig(format="%(asctime)-15s - %(message)s",
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.INFO)

    parser = argparse.ArgumentParser(description='CCC (Cheap and Cheerful Charger)')
    args = parser.parse_args()

    print('\n=== Cheap and Cheerful Charger ===\n')

    sys.stderr = sys.stdout

    if not IS_WINDOWS:
        atexit.register(turn_power_off)
        signal.signal(signal.SIGUSR1, handler)

    SingleInstanceThread().start()

    PowerControlThread().start()

    WatchdogThread().start()

    if False and IS_WINDOWS:
        listen_for_sleep()

    print("Main thread terminated.")


if __name__ == "__main__":

    main()
