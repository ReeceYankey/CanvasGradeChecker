# note: it is assumed you are using windows

# credit to getting version: https://stackoverflow.com/questions/57441421/how-can-i-get-chrome-browser-version-running-now-with-python
from win32com.client import Dispatch
from decouple import config
import os.path
import requests
import re
import zipfile
import io

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

def get_chrome_version():
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version = [get_version_via_com(p) for p in paths if p is not None][0]
    return version

def lookup_driver_version():
    major_version = re.search(r"\d.", get_chrome_version()).group(0)

    url = f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major_version}"
    print(f"fetching {url}")
    r = requests.get(url)
    if r.status_code != 200:
        raise Exception(f"There was an error connecting to {url}")
    
    return r.text

def install_chrome_driver():
    driver_version = lookup_driver_version()
    
    url = f"https://chromedriver.storage.googleapis.com/{driver_version}/chromedriver_win32.zip"
    print(f"fetching {url}")
    r = requests.get(url, stream=True)
    if r.status_code != 200:
        raise Exception(f"There was an error connecting to {url}")

    print("extracting zip file")
    download = r.content
    with zipfile.ZipFile(io.BytesIO(download)) as zip_ref:
        zip_ref.extractall()

    print("successfully installed chromedriver.exe")

def settings_file_exists():
    return os.path.isfile("settings.ini")

def driver_path_is_valid():
    return os.path.isfile(config("CHROME_DRIVER_PATH"))

def first_time_setup():
    # user input
    print("Initiating first time setup...")
    CHROME_DRIVER_PATH = input("Please enter the path of your chromedriver.exe (empty to auto-install to local directory):")
    CANVAS_USERNAME = input("Please enter your canvas username (optional):")
    CANVAS_PASSWORD = input("Please enter your canvas password (optional):")
    
    # install driver if needed
    if CHROME_DRIVER_PATH == "":
        install_chrome_driver()
        CHROME_DRIVER_PATH = "chromedriver.exe"

    # write settings.ini file
    settings = open("settings.ini", "w+")
    settings.write(("[settings]\n"
                    "CHROME_DRIVER_PATH={}\n"
                    "CANVAS_USERNAME={}\n"
                    "CANVAS_PASSWORD={}").format(CHROME_DRIVER_PATH, CANVAS_USERNAME, CANVAS_PASSWORD))
    settings.close()

def verify_configuration():
    if not settings_file_exists():
        print("settings.ini has not been detected.", end=' ')
        first_time_setup()
    if not driver_path_is_valid():
        raise Exception("Chrome driver path invalid")



if __name__ == '__main__':
    first_time_setup()