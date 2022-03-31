import platform     # For getting the operating system name
import subprocess   # For executing a shell command
import json         # For getting and parsing json
import time
import win32com.client as win32


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
    reads it's and returns the data
    :return: json
    """

    # Opening JSON file
    f = open(file)

    # returns JSON object as
    # a dictionary
    data = json.load(f)

    # Closing file
    f.close()

    return data



def get_message(file):
    """
    Opens file (str) if exist
    reads it's and returns the data
    :return: html
    """

    # Opening JSON file
    f = open(file)

    # Get the html
    data = f.read()

    # Closing file
    f.close()

    return data



def send_mail(data):
    """
    Sends the Mail to the address that is stored in data (json)
    Outlook is req to be running
    """

    # gets the recipient mail address from data
    recipient = data['recipient_email']

    # opens an outlook instance
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # crafts the mail
    mail.To = recipient
    mail.Subject = data['message_sub']
    mail.HTMLBody = get_message('message.html')

    mail.Send()



if __name__ == '__main__':

    # get data from file
    data = get_json("config.json")
    ip = data['ip']
    t = data['t_in_s']

    # checks if file died
    while True:

        if not ping(ip):
            send_mail(data)
            break

        time.sleep(t)


    input("Press Enter to continue...")