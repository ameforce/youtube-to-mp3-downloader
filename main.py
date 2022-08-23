from selenium.webdriver.common.by import By
from win10toast import ToastNotifier
from functools import partial
from selenium import webdriver
from selenium import common
from pathlib import Path
from winreg import *
import chromedriver_autoinstaller
import multiprocessing
import pyperclip
import requests
import time
import os


def input_url():
    print('다운로드 받을  Youtube 링크를 입력해주세요(Ctrl+V)')
    url = input('--> ')
    return url


def url_extract(url: str):
    refine_url = url[-11:]
    return refine_url


def url_recombination(url: str):
    return f'https://www.y2mate.com/kr/youtube-mp3/{url}'


def driver_check():
    driver_install_path = f'{os.getcwd()}\\lib\\chromedriver'
    chrome_ver = chromedriver_autoinstaller.get_chrome_version().split('.')[0]
    driver_path = f'{driver_install_path}\\{chrome_ver}\\chromedriver.exe'

    if os.path.exists(driver_install_path):
        print(f'chrome driver is installed: {driver_path}')
    else:
        print(f'chrome driver is not installed.')
        print(f'install the chrome driver(ver: {chrome_ver})')
        os.makedirs(driver_install_path)
        while not os.path.exists(driver_install_path):
            continue
        chromedriver_autoinstaller.install(False, driver_install_path)
    return driver_path


def refine_title(title: str):
    title = title.replace('/', '')
    title = title.replace('\\', '')
    title = title.replace(':', '')
    title = title.replace('*', '')
    title = title.replace('?', '')
    title = title.replace('"', '')
    title = title.replace('<', '')
    title = title.replace('>', '')
    title = title.replace('|', '')
    title = title.replace('.mp3 ', '')
    title = title.replace('.mp3', '')
    title = title.replace('.wav ', '')
    title = title.replace('.wav', '')
    return title


def extract_video_code(youtube_url: str):
    """
    :param youtube_url: A YouTube video link extracted from the clipboard.
                        Used to extract video code.
    :return:            Returns the video code extracted from youtube_url.
    :rtype:             str
    """
    return youtube_url[32:]


def redundancy_check(video_code: str):
    """
    :param video_code:  The last 11 digits of code extracted from youtube_url.
                        Use to identify videos.
    :return:            Returns True if a value that overlaps with video_code is detected in the registry.
                        Otherwise, return False.
    :rtype:             bool
    """

    read_path = r'SOFTWARE\\ENMSoft\\YoutubeToMP3Downloader'
    reg_handle = ConnectRegistry(None, HKEY_CURRENT_USER)
    key = OpenKey(reg_handle, read_path, 0, KEY_READ)
    i = 0
    while True:
        try:
            if video_code == EnumValue(key, i)[0]:
                return True
            else:
                i += 1
        except WindowsError:
            CloseKey(key)
            return False


def clipboard_monitoring():
    """
    :return:    If a YouTube video link is detected in the clipboard, and if redundancy_check is passed,
                a YouTube video link is returned.
                At this time, the YouTube video link is truncated from the 0th character to the 43rd character.
                It has the ability to refine the YouTube video link even if there is more text attached to it.
    :rtype:     str
    """

    clipboard = pyperclip.paste()
    youtube_url = None
    if 'https://www.youtube.com/watch?v=' in clipboard and len(clipboard) >= 43:
        if not redundancy_check(extract_video_code(clipboard[:43])):
            youtube_url = clipboard[:43]
    return youtube_url


def save_data(video_code: str, running_time: int, using_time: float):
    save_path = r'SOFTWARE\\ENMSoft\\YoutubeToMP3Downloader'
    CreateKey(HKEY_CURRENT_USER, save_path)

    reg_handle = ConnectRegistry(None, HKEY_CURRENT_USER)
    key = OpenKey(reg_handle, save_path, 0, KEY_WRITE)
    try:
        SetValueEx(key, video_code, 0, REG_SZ, f'{str(running_time)}, {str(using_time)}')
    except EnvironmentError:
        print('레지스트리 쓰기에 문제가 발생했습니다.')

    CloseKey(key)


def calc_time(running_time: int):
    read_path = r'SOFTWARE\\ENMSoft\\YoutubeToMP3Downloader'
    reg_handle = ConnectRegistry(None, HKEY_CURRENT_USER)
    key = OpenKey(reg_handle, read_path, 0, KEY_READ)
    result_running_time, result_using_time, i = 0, 0, 0
    while True:
        try:
            temp_time_data = EnumValue(key, i)[1].split(', ')
            temp_running_time, temp_using_time = float(temp_time_data[0]), float(temp_time_data[1])
            if temp_using_time != 0:
                result_running_time += temp_running_time
                result_using_time += temp_using_time
            i += 1
        except WindowsError:
            break
    CloseKey(key)
    if not result_running_time == 0 and not result_using_time == 0:
        # The download_rate_per_second variable means that you can download n video lengths per second (in seconds).
        download_rate_per_second = result_running_time / result_using_time
        return running_time / download_rate_per_second
    else:
        return 0


def download_runner(process_name: str, mp3_title: str, download_url: str, running_time: int, timeout: float):
    if process_name == 'downloader':
        start_time = time.time()

        file = requests.get(download_url, allow_redirects=True, timeout=timeout)
        open(f'{Path.home()}\\Downloads\\{mp3_title}.mp3', 'wb').write(file.content)

        using_time = time.time() - start_time

        return using_time
    elif process_name == 'counter':
        calculated_time = calc_time(running_time)
        toaster = ToastNotifier()
        while calculated_time > 0:
            if int(calculated_time % 1 * 10) == 0 and not calculated_time <= 1:
                toaster.show_toast(f'다운로드 알림', f'다운로드까지 {calculated_time:.1f}초 남았습니다.',
                                   'lib\\baby_yoda.ico', duration=1, threaded=True)
                # toaster.on_destroy(on_alert, f'다운로드까지 {calculated_time:.1f}초 남았습니다.')
                print(f'다운로드까지 {calculated_time:.1f}초 남았습니다.')
            calculated_time -= 0.1
            time.sleep(0.1)


def refine_duration(duration: str):
    """
    :param duration:    The running time of the YouTube video in the form of an unrefined string obtained through
                        find_element.
    :return:            Divide the unrefined {duration} into {hours}, {minutes}, {seconds}, respectively,
                        and convert it to int and return it.
    :rtype:             int, int, int
    """

    duration_time = duration[10:]
    refine_duration_hours = int(duration_time[0:2])
    refine_duration_minutes = int(duration_time[3:5])
    refine_duration_seconds = int(duration_time[6:8])
    return refine_duration_hours, refine_duration_minutes, refine_duration_seconds


def convert_duration_to_running_time(refine_duration_hours: int, refine_duration_minutes: int,
                                     refine_duration_seconds: int):
    """
    :param refine_duration_hours:       This corresponds to the {hours} of the portion obtained by refine_duration.
    :param refine_duration_minutes:     This corresponds to the {minutes} of the portion obtained by refine_duration.
    :param refine_duration_seconds:     This corresponds to the {seconds} of the portion obtained by refine_duration.
    :return:                            Returns the value of {duration} converted to seconds.
    :rtype:                             int
    """

    return refine_duration_hours * 3600 + refine_duration_minutes * 60 + refine_duration_seconds


def logic(youtube_url: str, chrome_driver: webdriver = None, recombined_url: str = None, timeout: float = 1.0,
          toaster: ToastNotifier = None):
    if youtube_url is not None:
        recombined_url = url_recombination(url_extract(youtube_url))
        print('connecting server...')
        toaster.show_toast(f'Download Started', f'Downloading to\n{Path.home()}\\Downloads',
                           'lib\\baby_yoda.ico', duration=3, threaded=True)

        chrome_driver.get(recombined_url)
        chrome_driver.implicitly_wait(2)

        # Download Button
        print('download data searching...')
        chrome_driver.find_element(by=By.XPATH, value='//*[@id="result"]/div/div[2]/div/div/div[2]').click()

        # Extract Running Time
        print('Extracting video running time...')
        duration = chrome_driver.find_element(by=By.XPATH, value='//*[@id="result"]/div/div[2]/div/p').text
        refine_duration_hours, refine_duration_minutes, refine_duration_seconds = refine_duration(duration)
        running_time = convert_duration_to_running_time(refine_duration_hours, refine_duration_minutes,
                                                        refine_duration_seconds)

        # Extract Download URL
        try:
            print('download url extracting...')
            download_url = chrome_driver.find_element(by=By.XPATH,
                                               value='//*[@id="process-result"]/div/a').get_attribute('href')
        except common.exceptions.NoSuchElementException:
            try:
                error_msg = chrome_driver.find_element(by=By.XPATH, value='//*[@id="process-result"]/div').text.splitlines()[1]
                save_data(extract_video_code(youtube_url), running_time, 0)
                print(error_msg)
                toaster = ToastNotifier()
                toaster.show_toast(f'Error', f'{error_msg}',
                                   'lib\\baby_yoda.ico', duration=5, threaded=True)
            except IndexError:
                print(f'url extracting index error: {timeout}')
                print(chrome_driver.find_element(by=By.XPATH, value='//*[@id="process-result"]/div').text.splitlines()[0])
                logic(youtube_url, chrome_driver, recombined_url, timeout, toaster)
                return
            except common.exceptions.NoSuchElementException:
                print(f'url extracting no such error: {timeout}')
                logic(youtube_url, chrome_driver, recombined_url, timeout, toaster)
                return

        # Extract MP3 Title
        print('mp3 title extracting...')
        mp3_title = refine_title(chrome_driver.find_element(by=By.XPATH, value='//*[@id="exampleModalLabel"]').text)

        # Download Start
        print(f'[ {mp3_title}.mp3 ] downloading...')
        try:
            download_runner_func = partial(download_runner, mp3_title=mp3_title, download_url=download_url,
                                           timeout=timeout, running_time=running_time)
            pool = multiprocessing.Pool(processes=2)
            using_time = pool.map(download_runner_func, ['downloader', 'counter'])[0]
            pool.close()
            print(f'[ {mp3_title}.mp3 ] download complete !!!')
            toaster = ToastNotifier()
            toaster.show_toast(f'Download Completed !!!', f'{mp3_title}\nhas been downloaded to\n{Path.home()}\\Downloads',
                               'lib\\baby_yoda.ico', duration=5, threaded=True)

            save_data(extract_video_code(youtube_url), running_time, using_time)
        except requests.exceptions.ConnectionError:
            print(f'connection error: {timeout}')
            logic(youtube_url, chrome_driver, recombined_url, timeout, toaster)
        except requests.exceptions.Timeout:
            print(f'timeout: {timeout}')
            logic(youtube_url, chrome_driver, recombined_url, timeout+0.1, toaster)


def init_system():
    chrome_driver_options = webdriver.ChromeOptions()
    chrome_driver_options.headless = True
    chrome_driver = webdriver.Chrome(driver_check(), chrome_options=chrome_driver_options)

    toaster = ToastNotifier()

    return chrome_driver, toaster


def main():
    chrome_driver, toaster = init_system()
    while True:
        logic(clipboard_monitoring(), chrome_driver=chrome_driver, toaster=toaster)
        time.sleep(0.1)


if __name__ == '__main__':
    main()
