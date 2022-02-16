import os
import sys
import shutil


def main():
    separator = os.path.sep
    path = sys.argv[1]
    name = path.split(separator)[-1]
    directory = path[:path.rfind(separator)]


    print(f"Original name: {name}")
    new_name = input("Please enter new name: ")

    if os.path.isfile(path):
        print("Using file mode")
        os.rename(path, directory + separator + new_name)
        print(f"File '{name}' renamed to '{new_name}'")
    else:
        print("Using folder mode")
        for item in os.listdir(path):
            if item.find(name) == -1:
                print(f"Item name {item} is incosistent! Exiting...")
                sys.exit(1)
        print(f"Creating new directory '{directory + separator + new_name}'")
        os.mkdir(directory + separator + new_name)
        for item in os.listdir(path):
            new_item = item.replace(name, new_name)
            print(f"Item '{item}' renamed to '{new_item}'")
            shutil.move(path + separator + item, directory + separator + new_name + separator + new_item)
        print(f"Deleting old directory '{path}'")
        os.rmdir(path)

    input("Press any key to exit")

if __name__ == '__main__':
    main()