import platform                         # For getting the operating system name
import subprocess                       # For executing a shell command
import json                             # For getting and parsing json
import time                             # For sleep
import win32com.client as win32         # For using outlook
import datetime                         # For getting logtime


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


def write_json(data, file):
    """
    writes data (json obj) to file (str)
    """

    with open(file, "w") as outfile:
        json.dump(data, outfile)


def read_file(file):
    """
    Opens file (str) if exist
    reads it's and returns the data
    :return: str
    """

    # Opening JSON file
    f = open(file)

    # Get the html
    data = f.read()

    # Closing file
    f.close()

    return data


def send_mail(data, online):
    """
    Sends the Mail to the address that is stored in data (json)
    Outlook has to be installed
    """

    if not sending:
        print("no mail was send. you are in debug mode!")
        return

    # gets the recipient mail address from data
    recipient = data['recipient_email']

    try:
        # opens an outlook instance
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        # crafts the mail
        mail.To = recipient
        mail.Subject = data['message_sub']

        if online:
            mail.HTMLBody = read_file('message_down.html')
        else:
            mail.HTMLBody = read_file('message_up.html')

        mail.Send()
        print('mail was sent to : ' + recipient)

    except:
        print('sending the mail to ' + recipient + ' failed')


def log(message):
    """
    logs a message (str) to the logfile with the current timestamp
    """

    time = str(datetime.datetime.now())

    f = open("log.txt", "a")

    f.write(time + ' ' + message + '\n')

    f.close()


if __name__ == '__main__':

    log('program started')

    version = '1.1'
    config_file = 'config.json'
    sending = True

    print("welcome to iIPi version " + version)

    while True:

        print("would you like to use the last settings?")
        print("(yes/no)")
        userinput = input()
        userinput.lower()

        if userinput == 'y' or userinput == 'yes':

            print('using old input')

            # get data from file

            data = get_json(config_file)
            ip = data['ip']
            t = data['t_in_s']
            r_mail = data['recipient_email']

            print('IP address:  ' + ip)
            print('test time:   ' + str(t))
            print('email:       ' + r_mail)

            break

        elif userinput == 'n' or userinput == 'no':

            print('enter new input')

            ip = input('enter IP: ')
            t = input('enter test time: ')
            t = int(t)
            r_mail = input('email: ')

            data = get_json(config_file)
            data['ip'] = ip
            data['t_in_s'] = t
            data['recipient_email'] = r_mail

            print('you want to safe the data to the config file?')
            userinput = input("type yes to save yes")
            userinput.lower()

            if userinput == 'y' or userinput == 'yes':
                print('saving input')
                write_json(data, config_file)
            else:
                print('not saving input')

            break

        else:

            print('invalid input')

    print('-------------------------------')
    print('starting pinging')
    print('-------------------------------')

    # checks if file died
    online = True
    while True:

        # when ping fails and server should be on
        if not ping(ip) and online:
            log('Server is down')
            send_mail(data, online)
            print('-------------------------------')
            print("Server is down")
            print('-------------------------------')
            online = False

        # when ping succeeds and server was down
        if ping(ip) and not online:
            log('Server is up')
            send_mail(data, online)
            print('-------------------------------')
            print("Server is up")
            print('-------------------------------')
            online = True

        time.sleep(t)
