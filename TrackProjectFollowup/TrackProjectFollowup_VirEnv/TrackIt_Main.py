import TrackFollowup as TF
from termcolor import colored
import colorama
import datetime as dt

colorama.init()
print(colored("\n\tTrackIT: TakeFollowUp() running...", "green"), "\n")

try:
    f = open('eLog.txt', 'a+')
    try:
        TF.TakeFollowUp()
    except Exception as e:
        f.write(str(dt.datetime.now()) + ":\n\t" + str(e) + "\n\n\n")
finally:
    f.close()

input("\nPress Enter to exit TrackIT...")
