import webbrowser
import speech_recognition as sr
import openai
import os
import win32com.client
import subprocess
import winreg
import re
import win32api
import googlesearch

import my_constants
from MyFaceRecognitionEngine import MyFaceRecognitionEngine

open_successful = True
openai.api_key = my_constants.OPEN_AI_KEY
user_name = "Unknown"


def say(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)


def sayUsingCommandLine(text):
    command = f"PowerShell -Command \"Add-Type -AssemblyName System.Speech; (New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{text}');\""
    os.system(command)


def foo(hive, flag):
    aReg = winreg.ConnectRegistry(None, hive)
    aKey = winreg.OpenKey(aReg, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                          0, winreg.KEY_READ | flag)

    count_subkey = winreg.QueryInfoKey(aKey)[0]

    software_list = []

    for i in range(count_subkey):
        software = {}
        try:
            asubkey_name = winreg.EnumKey(aKey, i)
            asubkey = winreg.OpenKey(aKey, asubkey_name)
            software['name'] = winreg.QueryValueEx(asubkey, "DisplayName")[0]

            try:
                software['version'] = winreg.QueryValueEx(asubkey, "DisplayVersion")[0]
            except EnvironmentError:
                software['version'] = 'undefined'

            try:
                software['location'] = winreg.QueryValueEx(asubkey, "InstallLocation")[0]
            except EnvironmentError:
                software['location'] = 'undefined'
            try:
                software['publisher'] = winreg.QueryValueEx(asubkey, "Publisher")[0]
            except EnvironmentError:
                software['publisher'] = 'undefined'
            software_list.append(software)
        except EnvironmentError:
            continue

    return software_list


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.record(source, 5)
        query = r.recognize_google(audio, language="en-in")
        return query


def find_file(root_folder, rex):
    for root, dirs, files in os.walk(root_folder):
        for f in files:
            result = rex.search(f)
            if result:
                print(os.path.join(root, f))
                break  # if you want to find only one


def find_file_in_all_drives(file_name):
    # create a regular expression for the file
    rex = re.compile(file_name)
    for drive in win32api.GetLogicalDriveStrings().split('\000')[:-1]:
        find_file(drive, rex)


def searchFile(rootDir, file_to_search):
    file_path = ""
    for relPath, dirs, files in os.walk(rootDir):
        files_in_lower = [x.lower() for x in files]
        if file_to_search in files_in_lower:
            index = files_in_lower.index(file_to_search)
            name = files[index]
            fullPath = os.path.join(rootDir, relPath, name)
            file_path = fullPath
            return file_path

    return file_path

def get_user_name():
    command = "whoami"
    result = subprocess.run(command, stdout=subprocess.PIPE).stdout.decode('utf-8')
    result_list = list(result.split("\\"))
    return result_list.pop().strip()


def openAapp(query):
    say("Searching App. Please Wait!")
    new_query = query.split("open")[1]
    result = askOpenAi(f"{new_query}.exe full file path in Windows 64")
    result_dic = dict(result)
    text = result_dic["text"]
    split_text = text.split("\\")
    for text in split_text:
        if ".exe" in text.lower():
            filename = text.lower().split(".exe")[0]
            filename_without_space = list(filename.lower().split(" ")).pop()
            file = f"{filename_without_space}.exe".lower()
            print(file)
            rootDir = "C:\\Program Files"
            rootDirX = "C:\\Program Files (x86)"
            rootUser = f"C:\\Users\{get_user_name()}"
            rootWindows = "C:\\Windows\System32"
            file_to_search = file
            filePath = searchFile(rootDir, file_to_search)
            if filePath != "":
                print(filePath)
                say("Opening App")
                subprocess.Popen(filePath.strip())
                break
            else:
                say("Looking for app in the system")
                filePath = searchFile(rootDirX, file_to_search)
                if filePath != "":
                    print(filePath)
                    say("Opening App")
                    subprocess.Popen(filePath.strip())
                    break
                else:
                    filePath = searchFile(rootUser, file_to_search)
                    if filePath != "":
                        print(filePath)
                        say("Opening App")
                        subprocess.Popen(filePath.strip())
                        break

                    else:
                        filePath = searchFile(rootWindows, file_to_search)
                        if filePath != "":
                            print(filePath)
                            say("Opening App")
                            subprocess.Popen(filePath.strip())
                            break
                        else:
                            say("Sorry application not found")
                            break
                        # check_app_in_system_and_open(query)
    return


def openApplication(install_location, appname, query):
    print(install_location)
    if install_location == None or install_location == "":
        say("Searching App. Please Wait!")
        new_query = query.split("open")[1]
        result = askOpenAi(f"{new_query}.exe full file path in Windows 64")
        result_dic = dict(result)
        text = result_dic["text"]
        split_text = text.split("\\")
        for text in split_text:
            if ".exe" in text.lower():
                filename = text.split(".exe")[0]
                file = f"{filename}.exe"
                print(file)
                rootDir = "C:\\Program Files"
                rootDirX = "C:\\Program Files (x86)"
                rootUser = "C:\\Users\ASUS"
                file_to_search = file
                filePath = searchFile(rootDir, file_to_search)
                if filePath != "":
                    print(filePath)
                    say("Opening App")
                    subprocess.Popen(filePath.strip())
                else:
                    filePath = searchFile(rootDirX, file_to_search)
                    if filePath != "":
                        print(filePath)
                        say("Opening App")
                        subprocess.Popen(filePath.strip())
                    else:
                        filePath = searchFile(rootUser, file_to_search)
                        if filePath != "":
                            print(filePath)
                            say("Opening App")
                            subprocess.Popen(filePath.strip())
                # os.chdir("C:\\")
                # res = subprocess.check_output(f"dir /s /b {filename}", shell=True, universal_newlines=True)
                # print(type(res))
                # print(res)

        return

    files = os.listdir(install_location)
    for file in files:
        if ".exe" in file and not "proxy" in file:
            app = file
            # app = f"{appname}.exe"
            say(f"Opening {appname}")
            command = f"{install_location}\\{app}"
            subprocess.Popen(command)
            return
    say("Sorry! Unable to open this app because of the security reasons.")


def searchGoogle():
    search = "What is Android Studio .exe file path"
    result = googlesearch.search(search, "com", lang='en', num=3, stop=3, pause=2)
    print(type(result))
    for r in result:
        print(r)


def askOpenAi(query):
    response = openai.Completion.create(
        engine='text-davinci-003',  # Use the ChatGPT model
        prompt=query,  # Your conversation prompt
        max_tokens=50  # Number of tokens for the response
    )
    print(response.choices[0])
    return response.choices[0]
    # for reply in response.choices:
    #     print(reply)


def find_all(name, path):
    result = []
    for root, dirs, files in os.walk(path):
        if name in files:
            result.append(os.path.join(root, name))
    return result


def listen_command_and_execute():
    global user_name
    while True:
        try:
            query = takeCommand()
            if "open" in query:
                if "paint" in query.lower():
                    subprocess.Popen("mspaint.exe")
                elif "notepad" in query.lower():
                    subprocess.Popen("notepad")
                elif "youtube" in query.lower():
                    say("Opening YouTube!")
                    webbrowser.open("www.youtube.com")
                elif "google" in query.lower() and "chrome" not in query.lower():
                    say("Opening Google!")
                    webbrowser.open("www.google.com")
                elif "facebook" in query.lower():
                    say("Opening Facebook!")
                    webbrowser.open("www.facebook.com")
                elif "wikipedia" in query.lower():
                    say("Opening Wikipedia!")
                    webbrowser.open("www.wikipedia.com")
                else:
                    check_app_in_system_and_open(query)

            elif ("who am i" in query.lower()) or ("who i am" in query.lower()) or ("my name" in query.lower()):
                if "Unknown" == user_name or None == user_name:
                    say("First let me see you with my eyes")
                    my_fr_engine = MyFaceRecognitionEngine()
                    user_name = my_fr_engine.get_user_name()
                    if "Unknown" == user_name or None == user_name:
                        say("Sorry! I am unable to recognize you. Please try again.")
                    else:
                        say(f"I can see that you are {user_name}")

                else:
                    say(f"Yes! I know! you are {user_name}!")


            elif "bye bye" in query:
                say("Bye! See You again.")
                break
        except Exception as e:
            print(Exception)


def check_app_in_system_and_open(query):
    split_query = query.split(" ")
    # appsLst = winapps.list_installed()
    appsLst = foo(winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY) + foo(winreg.HKEY_LOCAL_MACHINE,
                                                                           winreg.KEY_WOW64_64KEY) + foo(
        winreg.HKEY_CURRENT_USER, 0)

    query_set = set(split_query)
    for app in appsLst:
        name_set = set(app['name'].split(" "))
        if len(query_set.intersection(name_set)) > 0:
            file_name = (list)(query_set.intersection(name_set))
            openInstalledApp(app['location'], app['name'], query)
            return
    # If app is not listed in installed applications list. Search in whole system.
    openAapp(query)
    return

def openInstalledApp(install_location, appname, query):
    if len(install_location)>15:
        files = os.listdir(install_location)
        for file in files:
            if ".exe" in file and not "proxy" in file:
                app = file
                # app = f"{appname}.exe"
                say(f"Opening {appname}")
                command = f"{install_location}\\{app}"
                subprocess.Popen(command)
                return
        # If .exe file not found in the install location then call open app
        openAapp(query)
        return
    # If install location path is not found or is incorrect.
    else:
        openAapp(query)
        return

if __name__ == '__main__':
    say("Hello I am Manav. How can I help you")
    listen_command_and_execute()
