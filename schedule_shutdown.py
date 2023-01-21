import os
import time
from datetime import datetime
from datetime import timedelta

active_shutdown = True

#---------------------------

def unix_to_date(unix_value):
    return datetime.utcfromtimestamp(unix_value).strftime("%Y-%m-%d %H:%M:%S")

#---------------------------

def date_to_unix(date_string):
    return time.mktime(datetime.strptime(date_string, "%Y-%m-%d %H:%M:%S").utctimetuple())

#---------------------------

def timestamp_format(elapsedTime):
    mon, sec = divmod(elapsedTime, 60)
    hr, mon = divmod(mon, 60)
    elapsedTimeString = "%d:%02d:%02d" % (hr, mon, sec)
    return elapsedTimeString

#---------------------------

def check_time(target_time_str):
    global active_shutdown
    target_time_unix = date_to_unix(target_time_str)
    current_time_unix = time.time()
    delta = current_time_unix - target_time_unix

    print('Time:',unix_to_date(current_time_unix),
          '| Target:',target_time_str,
          '| ETA:',str(round(delta,3)),'s')

    t_bool = False
    if current_time_unix >= target_time_unix:
        print('t: -' + str(round(delta,3)) + 's')
        t_bool = True

    if delta > 3600:
        active_shutdown = False

    return t_bool

#---------------------------

def system_shutdown():
    print('system_shutdown in 60s')
    time.sleep(60)
    os.system('shutdown /s')

#---------------------------

#time format: "%Y-%m-%d %H:%M:%S"
shutdown_time = '2023-01-21 04:00:00'

#check frequency in minutes
check_freq_min = 5

#main loop
while True:
    timer = check_time(shutdown_time)
    if timer == True:
        break
    time.sleep(check_freq_min*60)

#safety-checked shutdown
if active_shutdown == True:
    system_shutdown()
else:
    print('Previus date used. Shutdown cancelled.')
