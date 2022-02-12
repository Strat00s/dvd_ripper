import os
import sys
import wmi
import ctypes
import argparse
from makemkv import MakeMKV
import keyboard
import logging
from reprint import output

logging.getLogger("makemkv").setLevel(logging.ERROR)

#  .\HandBrakeCLI.exe --preset-import-gui -v -i D:/VIDEO_TS/VTS_02_1.VOB -o movie.mp4 -e x264 -q 20 -B 160 -s "1,2,3,4,5,6" -a "1, 2, 3, 4, 5, 6"
#  .\HandBrakeCLI.exe -v -i D:/VIDEO_TS --min-duration 120 -t 2 -o test.mp4 --all-audio --all-subtitles


def getDrive(drive_letter):
    for cdrom in wmi.WMI().Win32_CDROMDrive():
        if (drive_letter == cdrom.Drive):
            return cdrom
    return None

def driveStatus(drive_letter):
    for cdrom in wmi.WMI().Win32_CDROMDrive():
        if (drive_letter == cdrom.Drive):
            return cdrom.MediaLoaded
    return None

def driveInfo(drive_letter):
    drive = getDrive(drive_letter)
    print(f"Drive '{drive.Id}' information:")
    print(f"  ID:           {drive.Id}")
    print(f"  Name:         {drive.Name}")
    print(f"  Manufacturer: {drive.Manufacturer}")
    print(f"  Status:       {drive.Status}")
    print(f"  Volume name:  {drive.VolumeName}")
    print(f"  Capabilities:")
    for i in range(0, len(drive.Capabilities)):
        print(f"    {drive.CapabilityDescriptions[i]}")

def listDrives(cdrom_letters):
    print("Available drives:")
    for letter in cdrom_letters:
        print(f"  {letter}, ", end="")
    print("")

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

#checks if file exists amd returns newly numberd name (without extension)
def numberFile(directory, title_name, extension, sub_directory="\\"):
    i = 1
    while os.path.exists(directory + "\\" + sub_directory + "\\" + title_name + extension):
        extension = "(" + str(i) + ").mkv"
        i += 1
    return title_name + extension[:-4]


def main():
    #arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("-l", "--list",       action="store_true", help="List available drives")
    parser.add_argument("-i", "--input",      type=str,            help="Input drive to use (eg. F:). If not set and you have only single drive, that one will be used")
    parser.add_argument("-d", "--directory",  type=str,            help="Directory where to rip the disc(s) (absolute path)")
    args = parser.parse_args()
    
    cdrom = ""
    directory = ""
    
    #driveInfo("D:")
    
    #get available drives
    cdrom_letters = []
    for cdrom in wmi.WMI().Win32_CDROMDrive():
        cdrom_letters.append(cdrom.Drive)
    
    
    #list drives only
    if args.list == True:
        listDrives(cdrom_letters)
        sys.exit(0)
    
    
    #check that user selected drive exists
    if args.input is not None:
        exists = False
        for letter in cdrom_letters:
            if (args.input == letter):
                exists = True
                print(f"Using CDROM drive '{args.input}'")
                cdrom = args.input
                break
        if not exists:
            print(f"CDROM drive '{args.input}' not found")
            listDrives(cdrom_letters)
            sys.exit(1)
    else:
        if (len(cdrom_letters) == 0):
            print("No CDROM drives found!")
            sys.exit(1)
        elif (len(cdrom_letters) == 1):
            print(f"Found one CDROM drive: '{cdrom_letters[0]}'")
            cdrom = cdrom_letters[0]
        else:
            print("Multiple CDROM drives found!")
            listDrives(cdrom_letters)
            print("Please specify which one to use using -i/--input")
            sys.exit(1)
    
    
    #get directory
    directory = os.getcwd()
    if args.directory is not None:
        directory = args.directory
    

    makemkv = MakeMKV(cdrom, progress_handler=showProgress, minlength=600)
    
    item_cnt = 0
    while True:
        #wait for disc to be inserted
        print("Please insert new/next disc (exit by pressing 'e')")
        status = None
        while status != True:
            status = driveStatus(cdrom)
            if keyboard.is_pressed("e"):
                print("Exiting")
                sys.exit(0)
        print(f'Disc inserted')

        #create title name
        title = getDrive(cdrom).VolumeName
        title = title.replace("_", " ")
        title = title.replace("-", " ")
        title = title.lower()
        tmp = title
        title = ""
        for word in tmp.split(" "):
            word = word[0].upper() + word[1:] + " "
            title += word
        title = title.strip()

        print(f"\nTitle:         {title}")

        #check item count
        item_cnt = len(os.listdir(directory))

        #dump all
        makemkv.mkv("all", directory)

        #we have more then one new item
        item_cnt = len(os.listdir(directory)) - item_cnt
        print(f"Number of new items: {item_cnt}")
        
        extension = ".mkv"
        if item_cnt > 1:
            title = title.strip(" 0123456789")
            sub_directory = title
            episode_cnt = 1
            if os.path.exists(sub_directory):
                episode_cnt = len(os.listdir(sub_directory)) + 1
                print(f"Folder '{title}' already exists. Continuing with part '{episode_cnt}'")
            else:
                print(f"Creating folder '{title}'")
                os.mkdir(sub_directory)
                print(f"Starting on part '{episode_cnt}'")
            print(f"Title:         {title} - díl xx.mkv")
            #print(f"Directory:     {directory}")
            print(f"Sub directory: {sub_directory}")

            for i in range(0, item_cnt):
                default_title = "title_t" + str(i).zfill(2) + extension
                title = sub_directory + " - díl " + str(episode_cnt).zfill(2)

                title = numberFile(directory, title, extension, sub_directory)
                tmp = sub_directory + "\\" + title + extension

                print(f"Rename and move {default_title} -> {tmp}")
                os.rename(directory + "\\" + default_title, directory + "\\" + tmp)
                episode_cnt += 1
        else:
            title = numberFile(directory, title, extension)
            print(f"Title:         {title}")
            print(f"Directory:     {directory}")
            os.rename(directory + "\\" + "title_t00" + extension, directory + "\\" + title + extension)
        
        ctypes.windll.WINMM.mciSendStringW(u"open " + cdrom + u" type cdaudio alias d_drive", None, 0, None)
        ctypes.windll.WINMM.mciSendStringW(u"set d_drive door open", None, 0, None)
        print("Done\n")

if __name__ == '__main__':
    main()