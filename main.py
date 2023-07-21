import multiprocessing
import sys
import driver
import app
import wmi

my_key = "Seagate Backup Plus Drive"
files = r"data"
c = wmi.WMI()


def check_for_key():
    for disk in c.Win32_LogicalDisk():
        if disk.VolumeName == my_key:
            print("ENCRYPTION DISK CONNECTED")
            return disk


def app_func():
    app.main()
    return app.selected_file.get()


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        multiprocessing.freeze_support()

    disk = check_for_key()
    if disk != None:
        try:
            filename = app_func()
        except KeyError:
            print("ERROR!!!")
            print("File does not have column 'LINKS'")
        except PermissionError:
            print("ERROR!!!")
            print("File Opened in background")
        except:
            print("ERROR!!!")

        try:
            driver.main(filename)
        except KeyError:
            print("ERROR!!!")
            print("File does not have column 'LINKS'")
        except PermissionError:
            print("ERROR!!!")
            print("File Opened in background")
        except FileNotFoundError:
            print("ERROR!!!")
            print("File Not Selected")
        except:
            print("ERROR!!!")

    else:
        print("ERROR!!!")
        print("ENCRYPTION DISK NOT CONNECTED")

    input("Press ENTER to quit...")
    sys.exit()
