#!/usr/bin/env python3

'''

v1 - 26 Jan 2018
    # Support for TP-Link HS100 charger
v2 - 31 Dec 2019
    # Refactoring, encapsulate logic in relay specific class

TODO
    # Ensure only one instance could run
    # run() instead of popen()
    # requests instead of urllib
    # 
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
#MIN_CHARGE, MAX_CHARGE = 45, 55
#MIN_CHARGE, MAX_CHARGE = 49, 51
MANUAL = 40, 60
MIN_CHARGE_MANUAL, MAX_CHARGE_MANUAL = MIN_CHARGE - 1, MAX_CHARGE + 1
MAX_ALERT_CHARGE = MAX_CHARGE + 5
MIN_ALERT_CHARGE = MIN_CHARGE - 5


@unique
class RelayState(Enum):
    ON = auto()
    OFF = auto()
    NA = auto()


def wifi_ssid() -> str:
    """Wifi ssid or empty string if not connected to wifi"""

    if not IS_WINDOWS:
        popen = subprocess.Popen(['iwgetid', '-r'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, _ = popen.communicate()
    else:
        popen = subprocess.Popen(['netsh', 'wlan', 'show', 'interfaces'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, _ = popen.communicate()

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


class Relay(ABC):

    @property
    @abstractmethod
    def state(self):
        return RelayState.NA

    @abstractmethod
    def turn_power(self, on_off):
        pass


class HIDRelay(Relay):

    @property
    def state(self):

        popen = subprocess.Popen(['hidusb-relay-cmd.exe', 'state'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        output, _ = popen.communicate()
        if 'R1=ON' in output.decode("utf-8"):
            return RelayState.ON
        elif 'R1=OFF' in output.decode("utf-8"):
            return RelayState.OFF
        else:
            return RelayState.NA

    def turn_power(self, on_off):

        if on_off:
            os.system('hidusb-relay-cmd.exe on 1')
        else:
            os.system('hidusb-relay-cmd.exe off 1')


class HS100Relay(Relay):

    @property
    def state(self):

        try:
            output = urllib.request.urlopen("http://192.168.1.103/status", timeout=TIMEOUT).read()
            data = json.loads(output)
            #print(output)
            if data['system']['get_sysinfo']['state'] is None:
                return RelayState.NA
            return RelayState.ON if data['system']['get_sysinfo']['state'] == 1 else RelayState.OFF
        except:
            print('Error: In HS100.state()')
            return RelayState.NA


    def turn_power(self, on_off):

        try:
            if on_off:
                urllib.request.urlopen("http://192.168.1.103/on", timeout=TIMEOUT)
            else:
                urllib.request.urlopen("http://192.168.1.103/off", timeout=TIMEOUT)
        except:
            print('Error: In HS100.turn_power()')


class EnergenieRelay(Relay):

    @property
    def state(self):
        return RelayState.NA

    def turn_power(self, on_off):

        if on_off:
            urllib.request.urlopen("http://192.168.1.108:8000/on", timeout=TIMEOUT)
        else:
            urllib.request.urlopen("http://192.168.1.108:8000/off", timeout=TIMEOUT)


class NoRelay(Relay):

    @property
    def state(self):
        return RelayState.NA

    def turn_power(self, on_off):
        pass


relay = EnergenieRelay()
#relay = NoRelay()


def power_plugged():

    return psutil.sensors_battery().power_plugged


def turn_power_off():

    try:
        logging.info('Program getting killed, turning power off now...')
        relay.turn_power(False)
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


def control(control=True):

    # Hack for manual charging
    if battery_percent() < MIN_CHARGE_MANUAL and not power_plugged():
        beep(1000, 1000)
    elif battery_percent() > MAX_CHARGE_MANUAL and power_plugged():
        beep(2000, 3000)

    logging.info(f'{battery_percent():.1f}% Relay={relay.state} Power={power_plugged()}')

    if not control:
        return

    if relay.state == RelayState.NA:
        #return
        pass

    if battery_percent() <= MIN_CHARGE:

        if relay.state == RelayState.OFF or relay.state == 'N/A':
            relay.turn_power(ON)
            logging.info(f'\t{battery_percent():.1f}% - Turn power ON')

            # Check power is indeed on - wait for a few secs to give the power the time to reach the
            # computer
            time.sleep(10)
            if power_plugged() != 'ON':
                logging.error('\t### Not charging although power turned ON')
                # beep(1000, 1000)  # Get rid of the pesky 2 beeps
                playsound('Battery_Alert.wav')

        # Turn power ON anyway to guard if the above command failed
        relay.turn_power(ON)

    elif battery_percent() >= MAX_CHARGE:

        if relay.state != 'OFF' or relay.state == RelayState.NA:
            relay.turn_power(True)
            logging.info(f'\t{battery_percent():.1f}% - Turn power OFF')

            # Check power is indeed off - wait a few secs
            time.sleep(10)
            if power_plugged():
                logging.error('\t### Relay stuck on ON position!?')
                #beep(1000, 1000)
                playsound('Battery_High.m4a')

        # Turn power OFF anyway to guard if the above command failed
        relay.turn_power(False)

    if relay.state == RelayState.OFF and power_plugged() == RelayState.ON:
        logging.warning('\t### Charging when not supposed to!?')
        beep(1000, 1000)

    if relay.state == RelayState.ON and not power_plugged():
        logging.warning('\t### Plug charger in or check why not charging!')
        beep(1000, 1000)


def test_on_off():

    while True:
        relay.turn_power(True)
        time.sleep(30)
        relay.turn_power(False)
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

    def __init__(self, control):

        threading.Thread.__init__(self)
        self.control = control

    def run(self):

        while True:

            try:
                control(self.control)
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
                logging.info(f'\t### Overcharged above {MAX_ALERT_CHARGE:.1f}% - {battery_percent():.1f}%')
                #if relay.state == 'ON':
                if psutil.sensors_battery().power_plugged:
                    beep(500, 3000)
                    time.sleep(0.1)
                    beep(500, 3000)
                    time.sleep(0.1)
                    beep(500, 3000)

            if battery_percent() <= MIN_ALERT_CHARGE:
                logging.info(f'\t### Undercharged below {MIN_ALERT_CHARGE:.1f}% - {battery_percent():.1f}%')
                #if relay.state == 'OFF':
                if not psutil.sensors_battery().power_plugged:
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


class SleepThread(threading.Thread):

    def __init__(self):

        threading.Thread.__init__(self)

    def run(self):

        SLEEP_AFTER_MINS = 5
        SLEEP_AFTER_SECS = SLEEP_AFTER_MINS * 60

        while True:

            popen = subprocess.Popen(['xprintidle'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            output, _ = popen.communicate()

            inactivity_secs = int(output)//1000

            print(f'inactivity_secs={inactivity_secs}')

            if inactivity_secs >= SLEEP_AFTER_SECS:

                print('Turning power off and going to sleep...')

                try:
                    relay.turn_power(OFF)
                except:
                    print('Error: Exception thrown when calling relay.turn_power(OFF)')
                
                time.sleep(5)
                os.system('systemctl suspend')

            else:

                print('Activity detected, not going to sleep!')

            time.sleep(SLEEP_AFTER_SECS + 10)



def main():

    print('\n=== Cheap and Cheerful Charger ===\n')

    logging.basicConfig(format="%(asctime)-15s - %(message)s",
                        datefmt='%Y-%m-%d %H:%M:%S',
                        level=logging.INFO)

    parser = argparse.ArgumentParser(description='CCC (Cheap and Cheerful Charger)')
    parser.add_argument('--nocontrol', help='no power control, just monitor', action='store_true')
    args = parser.parse_args()

    no_control = True if args.nocontrol else False
    control = not no_control

    if no_control:
        logging.info('Monitoring mode, power source not controlled')

    sys.stderr = sys.stdout

    if not IS_WINDOWS:
        atexit.register(turn_power_off)
        signal.signal(signal.SIGUSR1, handler)

    SingleInstanceThread().start()

    PowerControlThread(control).start()

    WatchdogThread().start()

    if control:
        SleepThread().start()

    if False and IS_WINDOWS:
        listen_for_sleep()


if __name__ == "__main__":

    main()
