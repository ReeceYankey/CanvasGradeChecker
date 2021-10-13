# from abc import ABC, abstractmethod
from sys import platform

if platform == 'win32':
    from win32com.client import Dispatch
from decouple import config
import os.path
import requests
import re
import zipfile
import io
import subprocess


# class GenericInstaller(ABC):
#     @classmethod
#     @abstractmethod
#     def get_chrome_version(cls):
#         pass

#     @classmethod
#     def lookup_chrome_driver_version(cls, chrome_version):
#         major_version = re.search(r'\d*(?=\.)', chrome_version).group(0)

#         url = f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major_version}"
#         print(f"fetching {url}")
#         r = requests.get(url)
#         if r.status_code != 200:
#             raise Exception(f"There was an error connecting to {url}")

#         return r.text

#     @classmethod
#     def install_driver(cls, url):
#         chrome_version = cls.get_chrome_version()
#         driver_version = cls.lookup_chrome_driver_version(chrome_version)
#         if platform == 'win32':
#             url = f"https://chromedriver.storage.googleapis.com/{driver_version}/chromedriver_win32.zip"
#         elif platform == 'linux':
#             url = f"https://chromedriver.storage.googleapis.com/{driver_version}/chromedriver_linux64.zip"
#         else:
#             raise Exception(f'Sorry, this platform ({platform}) is currently unsupported')

#         print(f"fetching {url}")
#         r = requests.get(url, stream=True)
#         if r.status_code != 200:
#             raise Exception(f"There was an error connecting to {url}")

#         print("extracting zip file")
#         download = r.content
#         with zipfile.ZipFile(io.BytesIO(download)) as zip_ref:
#             zip_ref.extractall()

#         print("successfully installed chromedriver")

# class LinuxInstaller(GenericInstaller):
#     def get_chrome_version(self):
#         completed_process = subprocess.run(['chromium', '--version'], check=True, capture_output=True, text=True)
#         version = re.search(r"\d*\.\d*\.\d*\.\d*", completed_process.stdout).group(0)
#         return version

# class WindowsInstaller(GenericInstaller):
#     @classmethod
#     def get_version_via_fs(cls, filename):
#         """get the version of chrome via the file system, given the path to chrome.exe"""
#         # credit to getting version: https://stackoverflow.com/questions/57441421/how-can-i-get-chrome-browser-version-running-now-with-python
#         parser = Dispatch("Scripting.FileSystemObject")
#         try:
#             version = parser.GetFileVersion(filename)
#         except Exception:
#             return None
#         return version

#     @classmethod
#     def get_chrome_version(cls):
#         paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
#                 r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
#         version = [cls.get_version_via_fs(p) for p in paths if p is not None][0]
#         return version

def lookup_chrome_driver_version(chrome_version):
    print(chrome_version)
    major_version = re.search(r'\d*(?=\.)', chrome_version).group(0)

    url = f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major_version}"
    print(f"fetching {url}")
    r = requests.get(url)
    if r.status_code != 200:
        raise Exception(f"There was an error connecting to {url}")

    return r.text


driver_url_LUT = {
    "linux": {
        "chrome": "https://chromedriver.storage.googleapis.com/{}/chromedriver_linux64.zip"
    },
    "win32": {
        "chrome": "https://chromedriver.storage.googleapis.com/{}/chromedriver_win32.zip"
    }
}


def get_driver_url(browser, driver_version):
    return driver_url_LUT[platform][browser].format(driver_version)


def download_driver(browser, driver_version):
    url = get_driver_url(browser, driver_version)

    print(f"fetching {url}")
    r = requests.get(url, stream=True)
    if r.status_code != 200:
        raise Exception(f"There was an error connecting to {url}")

    print("extracting zip file")
    download = r.content
    with zipfile.ZipFile(io.BytesIO(download)) as zip_ref:
        zip_ref.extractall()

    print("successfully installed chromedriver")


# ---------------------------------------------------------------------------------------------------------------------
def linux_get_chrome_version():
    completed_process = subprocess.run(['chromium', '--version'], check=True, capture_output=True, text=True)
    version = re.search(r"\d*\.\d*\.\d*\.\d*", completed_process.stdout).group(0)
    return version


def win32_get_version_via_fs(filename):
    """get the version of chrome via the file system, given the path to chrome.exe"""
    # credit to getting version: https://stackoverflow.com/questions/57441421/how-can-i-get-chrome-browser-version-running-now-with-python
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        print("error")
        return None
    print(f"version {version}")
    return version


def win32_get_chrome_version():
    paths = [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
             r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]
    version= list(filter(None, [win32_get_version_via_fs(p) for p in paths]))[0]
    print(f"got version {version}")
    return version


# ---------------------------------------------------------------------------------------------------------------------

func_get_chrome_version_LUT = {
    "linux": {
        "chrome": linux_get_chrome_version
    },
    "win32": {
        "chrome": win32_get_chrome_version
    }
}
driver_filename_LUT = {
    "linux": {
        "chrome": "chromedriver",
        "firefox": "geckodriver"  # FIXME doesn't work as a path
    },
    "win32": {
        "chrome": "chromedriver.exe",
        "firefox": "geckodriver.exe"
    }
}


# ---------------------------------------------------------------------------------------------------------------------
def install_driver(browser):
    get_chrome_version = func_get_chrome_version_LUT[platform][browser]
    print(platform, browser, get_chrome_version)
    version = get_chrome_version()
    print(version)
    driver_version = lookup_chrome_driver_version(version)
    # url = get_driver_url(browser, driver_version)
    download_driver(browser, driver_version)
    return driver_filename_LUT[platform]['chrome']


def settings_file_exists():
    return os.path.isfile("settings.ini")


def driver_path_is_valid():
    return os.path.isfile(config("DRIVER_PATH"))


def first_time_setup():
    # user input
    print("Initiating first time setup...")

    while (True):
        try:
            inp = int(input("Please enter the browser you would like selenium to use (1: chrome, 2: firefox):"))
            if inp == 1:
                PREFERRED_BROWSER = 'chrome'
                break
            elif inp == 2:
                PREFERRED_BROWSER = 'firefox'
                break
        except ValueError:
            pass

    # handle chrome driver if using chrome
    if PREFERRED_BROWSER == 'chrome':
        DRIVER_PATH = input("Please enter the path of your chromedriver (empty to auto-install to local directory):")
        if DRIVER_PATH == '':
            DRIVER_PATH = install_driver(PREFERRED_BROWSER)
    elif PREFERRED_BROWSER == 'firefox':
        DRIVER_PATH = input(
            "Please enter the full path (/full/path/to/geckodriver) of your geckodriver (downloadable at https://github.com/mozilla/geckodriver/releases/tag/latest):")

    CANVAS_USERNAME = input("Please enter your canvas username (optional):")
    CANVAS_PASSWORD = input("Please enter your canvas password (optional):")

    # write settings.ini file
    settings = open("settings.ini", "w+")
    settings.write(("[settings]\n"
                    "PREFERRED_BROWSER={}\n"
                    "DRIVER_PATH={}\n"
                    "CANVAS_USERNAME={}\n"
                    "CANVAS_PASSWORD={}").format(PREFERRED_BROWSER, DRIVER_PATH, CANVAS_USERNAME, CANVAS_PASSWORD))
    settings.close()


def verify_configuration():
    if not settings_file_exists():
        print("settings.ini has not been detected.", end=' ')
        first_time_setup()
    if not driver_path_is_valid():
        raise Exception("Chrome driver path invalid")


if __name__ == '__main__':
    first_time_setup()
