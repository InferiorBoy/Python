#!/usr/bin/env python

import os
import wmi
import time
import socket
import base64
import smtplib
import logging
from email.mime.text import MIMEText


def GetSrv(designation):
    """Get stopped service name and caption,
    Filtration 'designation' service whether there is 'Stopped'.

    :return: service state
    """
    c = wmi.WMI()
    ret = dict()
    for service in c.Win32_Service():
        state, caption = service.State, service.Caption
        if state == 'Stopped':
            t = ret.get(state, [])
            t.append(caption)
            ret[state] = t
    # If 'designation' service in the 'Stopped', return status is 'down'
    if designation in ret.get('Stopped'):
        logging.error('Service [%s] is down, try to restart the service. \r\n' % designation)
        return 'down'
    return True


def Monitor(sname):
    """Send the machine IP port 20000 socket request,
    If capture the abnormal returns the string 'ex'.

    :return: string 'ex'
    """
    s = socket.socket()
    s.settimeout(3)  # timeout
    host = ('127.0.0.1', 20000)
    try:  # Try connection to the host
        s.connect(host)
    except socket.error as e:
        logging.warning('[%s] service connection failed: %s \r\n' % (sname, e))
        return 'ex'
    return True


def RestartSocket(rstname, conn, run):
    """First check whether the service is stopped,
    if stop, start the service directly.
    The check whether the zombies,
    if a zombie, then restart the service.

    :return: flag or True
    """
    flag = False
    try:
        # From GetSrv() to obtain the return value, the return value
        if run == 'down':
            ret = os.system('sc start "%s"' % rstname)
            if ret != 0:
                raise Exception('[Errno %s]' % ret)
            flag = True
        elif conn == 'ex':
            retStop = os.system('sc stop "%s"' % rstname)
            retSart = os.system('sc start "%s"' % rstname)
            if retSart != 0:
                raise Exception('retStop [Status code %s] '
                                'retSart [Status code %s] ' % (retStop, retSart))
            flag = True
        else:
            logging.info('[%s] service running status to normal' % rstname)
            return True
    except Exception as e:
        logging.warning('[%s] service restart failed: %s \r\n' % (rstname, e))
        return flag


def SendMail(to_list, sub, contents):
    """Send alarm mail.

    :return: flag
    """
    mail_server = 'mail.stmp.com'  # STMP Server
    mail_user = 'YouAccount'  # Mail account
    mail_pass = base64.b64decode('Password')  # The encrypted password
    mail_postfix = 'smtp.com'  # Domain name

    me = 'Monitor alarm<%s@%s>' % (mail_user, mail_postfix)
    message = MIMEText(contents, _subtype='html', _charset='utf-8')

    message['Subject'] = sub
    message['From'] = me
    message['To'] = ';'.join(to_list)

    flag = False  # To determine whether a mail sent successfully
    try:
        s = smtplib.SMTP()
        s.connect(mail_server)
        s.login(mail_user, mail_pass)
        s.sendmail(me, to_list, message.as_string())
        s.close()
        flag = True
    except Exception, e:
        logging.warning('Send mail failed, exception: [%s]. \r\n' % e)

    return flag


def main(sname):
    """Parameter type in the name of the service need to monitor,
    perform functions defined in turn, and the return value is correct.
    After the program is running, will test three times,
    each time interval to 10 seconds.

    :return: retValue
    """
    retry = 3
    count = 0
    retValue = False  # Used return to the state of the socket
    while count < retry:
        ret = Monitor(sname)
        if ret != 'ex':  # If socket connection is normaol, return retValue
            retValue = ret
            return retValue
        isDown = GetSrv(sname)
        RestartSocket(rstname=sname, conn=ret, run=isDown)

        host = socket.gethostname()
        address = socket.gethostbyname(host)
        mailto_list = ['mail@smtp.com', ]  # Alarm contacts
        SendMail(mailto_list, 'Alarm',
                 ' <h4>Level: <u>ERROR</u></br> Host name: %s</br>'
                 ' IP Address: %s</br>'
                 ' Service name:</h4> <h5>%s</h5>'
                 % (host, address, sname))
        count += 1
        time.sleep(10)
    else:
        logging.error('[%s] service try to restart more than three times \r\n' % sname)

    return retValue


if __name__ == '__main__':

    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(message)s',
                        datefmt='%Y/%m/%d %H:%M:%S',
                        filename='D:\\logs\\Monitor.log',
                        filemode='ab')

    name = 'ServiceName'
    response = main(name)
    if response:
        logging.info('The [%s] service connection is normal \r\n' % name)
