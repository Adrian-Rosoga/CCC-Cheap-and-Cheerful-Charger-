import json
import socket
import logging
from struct import pack
from switch_plugins.switch import Switch

IP_PORT = '192.168.1.157:9999'
TIMEOUT = 10


def encrypt(string):
    key = 171
    result = pack('>I', len(string))
    for i in string:
        a = key ^ ord(i)
        key = a
        result += bytes([a])
    return result


def decrypt(string):
    key = 171
    result = ""
    for i in string:
        a = key ^ i
        key = i
        result += chr(a)
    return result


COMMANDS = {'info'     : '{"system":{"get_sysinfo":{}}}',
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


def sendCommand(cmd):
    ip, port = IP_PORT.split(':')
    port = int(port)
    try:
        sock_tcp = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock_tcp.settimeout(TIMEOUT)
        sock_tcp.connect((ip, port))
        sock_tcp.settimeout(None)
        sock_tcp.send(encrypt(cmd))
        data = sock_tcp.recv(2048)
        sock_tcp.close()
        return data
    except socket.error:
        logging.error(f'Connect failure to smartplug at {IP_PORT}')


class HS100Switch(Switch):

    @property
    def state(self):
        response = sendCommand(COMMANDS['info'])
        if response is None:    # Because on holiday for example
            return Switch.State.NA
        info = decrypt(response)
        info = '{' + info[5:]
        data = json.loads(info)
        return Switch.State.ON if data['system']['get_sysinfo']['relay_state'] == 1 else Switch.State.OFF

    def turn_on(self):
        sendCommand(COMMANDS['on'])

    def turn_off(self):
        sendCommand(COMMANDS['off'])
