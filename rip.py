import os
import sys
import wmi
import ctypes
import argparse
from makemkv import MakeMKV, ProgressParser
from pyparsedvd import load_vts_pgci
import time
import keyboard

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


#TODO read volume name using Win32_CDROMDrive
#if multiple items would be dumped
    #TODO if directory exists -> count number of items in directory and continue from last item
    #TODO dump all
    #TODO count number of new items -> compare with old count -> get difference
    #TODO rename them according to the difference
#if single item will be dumped
    #TODO if item exists -> add windows naming '(x)', starting from '1'
    #todo rename 'title_t00.mkv' to 'volume_name.mkv'
#TODO reset counters and wait for new disc


def continuous(cdrom, directory):
    with ProgressParser() as progress:
        makemkv = MakeMKV(cdrom, progress_handler=progress.parse_progress, minlength=600)
        while True:
            #wait for disc to be inserted
            print("Please insert new disc (exit by pressing 'e')")
            status = None
            while status != True:
                status = driveStatus(cdrom)
                if keyboard.is_pressed("e"):
                    print("Exiting")
                    sys.exit(0)
            print(f'Disc inserted')
            volume_name = getDrive(cdrom).VolumeName
            

            #create title name
            if args.output is not None:
                title = args.output
            else:
                title = disc_info["disc"]["name"]
                title = title.replace("_", " ")
                title = title.replace("-", " ")
                title = title.lower()
                tmp = title
                title = ""
                for word in tmp.split(" "):
                    word = word[0].upper() + word[1:] + " "
                    title += word
                title = title.strip()

            multidiscs = False
            #name file and folder acordingly
            if args.count > 1 or disc_info["title_count"] > 1:
                multidiscs = True
                title = title.strip(" 0123456789")
                #print(f"A folder '{title}' needs to be created")
                directory = directory + "\\" + title
                if os.path.exists(directory):
                    if args.continuous:
                        dir_len = len(os.listdir(directory))
                        print(f"Folder already exists and containes '{dir_len}' items")
                        episode_cnt = dir_len + 1
                    else:
                        print("Folder already exists. Skipping creation")
                else:
                    print(f"Creating folder '{title}'")
                    os.mkdir(directory)
            #    title = title + " - díl"
            #else:
            #    title += ".mkv"

            print(f"Directory: {directory}")
            print(f"Title: {title}")

            episode_cnt = 1
            next_disc = 1

            while True:
                out = makemkv.mkv("all", directory)
                print(out)
                if multidiscs:
                    for i in range(0, disc_info["title_count"]):
                        os.rename(directory + "\\" + "title_t" + str(i).zfill(2) + ".mkv", directory + "\\" + title + " - díl " + str(episode_cnt).zfill(2) + ".mkv")
                        episode_cnt += 1
                else:
                    os.rename(directory + "\\" + "title_t00.mkv", directory + "\\" + title + ".mkv")

                ctypes.windll.WINMM.mciSendStringW(u"open " + cdrom + u" type cdaudio alias d_drive", None, 0, None)
                ctypes.windll.WINMM.mciSendStringW(u"set d_drive door open", None, 0, None)

                if args.count - next_disc < 1:
                    print("Done")
                    sys.exit(0)
                next_disc += 1

                print(f"Please insert disc {next_disc}/{args.count}")
                print("If you wish to skip this disc, press 's'")

                #wait for disc to be inserted or skip
                status = None
                while status != True:
                    status = driveStatus(cdrom)
                    if keyboard.is_pressed("e"):
                        print("Exiting")
                        break
                if status != True:
                    episode_cnt += int(input("How many episodes to skip?: "))
                    print(f"Next episode will have number '{episode_cnt}'")

                print("Disc inserted")


def single(cdrom, directory, args):
    with ProgressParser() as progress:
        makemkv = MakeMKV(cdrom, progress_handler=progress.parse_progress, minlength=600)
        print(f'Disc inserted')
        #disc_info = makemkv.info()

        #create title name
        if args.output is not None:
            title = args.output
        else:
            title = disc_info["disc"]["name"]
            title = title.replace("_", " ")
            title = title.replace("-", " ")
            title = title.lower()
            tmp = title
            title = ""
            for word in tmp.split(" "):
                word = word[0].upper() + word[1:] + " "
                title += word
            title = title.strip()

        multidiscs = False
        #name file and folder acordingly
        if args.count > 1 or disc_info["title_count"] > 1:
            multidiscs = True
            title = title.strip(" 0123456789")
            #print(f"A folder '{title}' needs to be created")
            directory = directory + "\\" + title
            if os.path.exists(directory):
                if args.continuous:
                    dir_len = len(os.listdir(directory))
                    print(f"Folder already exists and containes '{dir_len}' items")
                    episode_cnt = dir_len + 1
                else:
                    print("Folder already exists. Skipping creation")
            else:
                print(f"Creating folder '{title}'")
                os.mkdir(directory)
        #    title = title + " - díl"
        #else:
        #    title += ".mkv"

        print(f"Directory: {directory}")
        print(f"Title: {title}")

        episode_cnt = 1
        next_disc = 1

        while True:
            out = makemkv.mkv("all", directory)
            print(out)
            if multidiscs:
                for i in range(0, disc_info["title_count"]):
                    os.rename(directory + "\\" + "title_t" + str(i).zfill(2) + ".mkv", directory + "\\" + title + " - díl " + str(episode_cnt).zfill(2) + ".mkv")
                    episode_cnt += 1
            else:
                os.rename(directory + "\\" + "title_t00.mkv", directory + "\\" + title + ".mkv")

            ctypes.windll.WINMM.mciSendStringW(u"open " + cdrom + u" type cdaudio alias d_drive", None, 0, None)
            ctypes.windll.WINMM.mciSendStringW(u"set d_drive door open", None, 0, None)

            if args.count - next_disc < 1:
                print("Done")
                sys.exit(0)
            next_disc += 1

            print(f"Please insert disc {next_disc}/{args.count}")
            print("If you wish to skip this disc, press 's'")

            #wait for disc to be inserted or skip
            status = None
            while status != True:
                status = driveStatus(cdrom)
                if keyboard.is_pressed("s"):
                    print("Skipping")
                    break
            if status != True:
                episode_cnt += int(input("How many episodes to skip?: "))
                print(f"Next episode will have number '{episode_cnt}'")

            print("Disc inserted")


def main():
    #arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("-l", "--list",       action="store_true", help="List available drives")
    parser.add_argument("-i", "--input",      type=str,            help="Input drive to use (eg. F:). If not set and you have only single drive, that one will be used")
    parser.add_argument("-o", "--output",     type=str,            help="Output file name. If not set, the program will try to guess the best name")
    parser.add_argument("-c", "--count",      type=int, default=1, help="Use multiple discs (all files will be dumped to directory named after the first disc)")
    parser.add_argument("-C", "--continuous", action="store_true", help="Continuously rip discs one after the other (exit by pressing 'e' while waiting for new drive). Does not support multidisc!")
    parser.add_argument("-d", "--directory",  type=str,            help="Directory where to rip the disc(s) (absolute path)")
    args = parser.parse_args()
    
    cdrom = ""
    output_path = ""
    count = 0
    directory = ""
    
    driveInfo("D:")
    
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
    
    if args.continuous:
        continuous(cdrom, directory)
    else:
        single(cdrom, directory, args)
    sys.exit(0)
    
    #wait for disc to be inserted
    last_status = None
    while True:
        status = driveStatus(cdrom)
        if last_status != status:
            last_status = status
            if last_status == False:
                print("Please insert disc")
            elif last_status == True:
                break
            else:
                print(f"Status returned {last_status}")
                sys.exit(1)
    
    title = ""
    #start rip
    with ProgressParser() as progress:
        makemkv = MakeMKV(cdrom, progress_handler=progress.parse_progress, minlength=600)
        while True:
            print(f'Disc inserted')
            #disc_info = makemkv.info()

            #create title name
            if args.output is not None:
                title = args.output
            else:
                title = disc_info["disc"]["name"]
                title = title.replace("_", " ")
                title = title.replace("-", " ")
                title = title.lower()
                tmp = title
                title = ""
                for word in tmp.split(" "):
                    word = word[0].upper() + word[1:] + " "
                    title += word
                title = title.strip()

            multidiscs = False
            #name file and folder acordingly
            if args.count > 1 or disc_info["title_count"] > 1:
                multidiscs = True
                title = title.strip(" 0123456789")
                #print(f"A folder '{title}' needs to be created")
                directory = directory + "\\" + title
                if os.path.exists(directory):
                    if args.continuous:
                        dir_len = len(os.listdir(directory))
                        print(f"Folder already exists and containes '{dir_len}' items")
                        episode_cnt = dir_len + 1
                    else:
                        print("Folder already exists. Skipping creation")
                else:
                    print(f"Creating folder '{title}'")
                    os.mkdir(directory)
            #    title = title + " - díl"
            #else:
            #    title += ".mkv"

            print(f"Directory: {directory}")
            print(f"Title: {title}")

            episode_cnt = 1
            next_disc = 1

            while True:
                out = makemkv.mkv("all", directory)
                print(out)
                if multidiscs:
                    for i in range(0, disc_info["title_count"]):
                        os.rename(directory + "\\" + "title_t" + str(i).zfill(2) + ".mkv", directory + "\\" + title + " - díl " + str(episode_cnt).zfill(2) + ".mkv")
                        episode_cnt += 1
                else:
                    os.rename(directory + "\\" + "title_t00.mkv", directory + "\\" + title + ".mkv")

                ctypes.windll.WINMM.mciSendStringW(u"open " + cdrom + u" type cdaudio alias d_drive", None, 0, None)
                ctypes.windll.WINMM.mciSendStringW(u"set d_drive door open", None, 0, None)

                if args.count - next_disc < 1:
                    print("Done")
                    sys.exit(0)
                next_disc += 1

                print(f"Please insert disc {next_disc}/{args.count}")
                print("If you wish to skip this disc, press 's'")

                #wait for disc to be inserted or skip
                status = None
                while status != True:
                    status = driveStatus(cdrom)
                    if keyboard.is_pressed("s"):
                        print("Skipping")
                        break
                if status != True:
                    episode_cnt += int(input("How many episodes to skip?: "))
                    print(f"Next episode will have number '{episode_cnt}'")

                print("Disc inserted")
    
        #if args.count is not None:
    
        #makemkv.mkv(i, "O:/This PC/Documents/0_Projects/dvd_ripper/")
    
        #print(title)
        #makemkv.mkv(0, "O:/This PC/Documents/0_Projects/dvd_ripper/")
        #for i in range(0, disc_info["title_count"]):
        #    print(f"Title {i}") 
        #    makemkv.mkv(i, "O:/This PC/Documents/0_Projects/dvd_ripper/")
    exit(0)

if __name__ == '__main__':
    main()