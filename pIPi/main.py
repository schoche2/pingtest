import platform     # For getting the operating system name
import subprocess   # For executing a shell command
import json         # For getting and parsing json
import smtplib      # For sending Mail
from email.message import EmailMessage
import time         # For sleep


def ping(host):
    """
    Returns True if host (str) responds to a ping request.
    Remember that a host may not respond to a ping (ICMP) request even if the host name is valid.
    """

    # Option for the number of packets as a function of
    param = '-n' if platform.system().lower() == 'windows' else '-c'

    # Building the command. Ex: "ping -c 1 google.com"
    command = ['ping', param, '1', host]

    return subprocess.call(command) == 0



def get_json(file):
    """
    Opens file (str) if exist
    reads it's and returns the text
    :return: str
    """

    # Opening JSON file
    f = open(file)

    # returns JSON object as
    # a dictionary
    data = json.load(f)

    # Closing file
    f.close()

    return data


def send_mail(data):

    sender = data['sender_email']
    password = data['sender_password']
    recipient = data['recipient_email']
    host = data['smpt_host']
    port = data['smpt_port']

    msg_body = 'Email sent using outlook!'

    # action
    msg = EmailMessage()
    msg['subject'] = 'Email sent using outlook.'
    msg['from'] = sender
    msg['to'] = recipient
    msg.set_content(msg_body)

    with smtplib.SMTP_SSL(host, port) as smtp:
        smtp.login(sender, password)

        smtp.send_message(msg)


if __name__ == '__main__':

    data = get_json("config.json")
    ip = data['ip']
    t = data['t_in_s']
    """
    while True:
        time.sleep(t)
        if not ping(ip):
            send_mail(data)
            break
    """
    send_mail(data)

