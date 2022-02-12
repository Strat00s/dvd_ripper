import os
from sys import exec_prefix
import time

#TODO create class from this
def drawBar(current, max, min=0, padding_l=0, padding_r=0, point="#", decimal=False, percent=False, time=False):
    #TODO measure time
    bar_width = os.get_terminal_size().columns - padding_l - padding_r

    if percent:
        bar_width -= 5
    decimal_bar = 0

    current_bar = int(current * ((bar_width - 2) / max))

    #create bar
    bar = "["
    for i in range(min, current_bar):
        bar += point

    #print decimal part
    decimal_bar = int(current * ((bar_width - 2) / (max / 10)))
    decimal_bar = decimal_bar - current_bar * 10
    if decimal and decimal_bar != 0:
        bar += str(decimal_bar)
    bar = bar.ljust(bar_width - 1)
    bar += "]"

    if percent:
        bar += " " + str(int(current * (100 / max))).rjust(3) + "%"

    print(bar, end="")

def showProgress(task_description: str, progress: int, max: int):
    #format task_description
    description_len = 28
    task_description += ":"

    if len(task_description) > description_len:
        description_len = len(task_description)
    
    task_description = task_description.ljust(description_len)
    
    if showProgress.last_progress > progress or showProgress.last_description != task_description:
        #finish last task
        if showProgress.last_progress < max:
            print(end="\r")
            print(showProgress.last_description, end=" ")
            drawBar(max, max, padding_l = description_len + 1, percent=True, decimal=True)
        
        showProgress.last_description = task_description

        #print current task at start
        showProgress.last_progress = progress
        print(task_description, end=" ")
        drawBar(progress, max, padding_l = description_len + 1, percent=True, decimal=True)
        print("", end="\r")
    else:
        showProgress.last_progress = progress
        print("", end="\r")
        print(task_description, end=" ")
        drawBar(progress, max, padding_l = description_len + 1, percent=True, decimal=True)
        showProgress.last_max = max
showProgress.last_description = ""        
showProgress.last_progress = 9999999999999999999
showProgress.last_max = 0

def main():
    #extension = ".mkv"
    #tmp = "awfunwiufnoqwndnqwodqw"
    #i = 1
    #while i < 11:
    #    print(tmp + extension)
    #    extension = "(" + str(i) + ").mkv"
    #    time.sleep(0.5)
    #    i += 1
    max = 65536
    step = 25
    i = 0
    while i < 100 + 1:
        showProgress("Scanning CD-ROM devices", i, 100)
        time.sleep(0.5)
        i += 1
    i = 0
    while i < max - max / 4:
        showProgress("Processing title sets", i, max)
        #time.sleep(0.01)
        i += 1
    i = 0
    while i < max - max / 2:
        showProgress("Processing titles", i, max)
        #time.sleep(0.01)
        i += 1
    i = 0
    while i < max - max / 3:
        showProgress("Decrypting data", i, max)
        #time.sleep(0.01)
        i += 1
    i = 0
    while i < max + 1:
        showProgress("Analyzing seamless segments", i, max)
        #time.sleep(0.01)
        i += 1
    i = 0
    while i < max + 1:
        showProgress("Saving to MKV file", i, max)
        #time.sleep(0.0001)
        i += 1
    i = 0
    

if __name__ == '__main__':
    main()
