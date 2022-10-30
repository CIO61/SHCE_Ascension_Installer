import os
import subprocess as sp
import sys
import traceback

import colorama
import pymsgbox
import requests
import shutil
import zipfile
import hashlib

from colorama import Fore
# noinspection PyPackageRequirements
from win32com import client

original_maps_set = {
    "Best_Friends.map", "Bird_in_flight.map", "Coastal_Trap.map", "Crossroads.map", "Divided.map", "Enclosure.map",
    "Fury_Bay.map", "Jealous_Neighbours.map", "MP-No mans land.map", "MP-No where to hide.map", "MP-Slopes of doom.map",
    "Lionheart.map", "Look_out.map", "MP-Divided.map", "MP-Downhill Scrum.map", "MP-Snake River.map",
    "MP-Surrounded.map", "MP-The rocky divide.map", "MP-Two Falls.map", "MP-Valley of the lords.map",
    "Phoenix.map", "Rivers_Fork.map", "Snake_River.map", "Spider_Island.map", "Swampy_Island.map",
    "The_Host.map", "Three_Little_Pigs.map", "Ultimate_Victory.map", "Wazirs_Fortress.map", "Wide_Open_Plain.map"
}
game_path = os.path.abspath("..\\")


def fc(color, text):
    return f'{Fore.__dict__.get(color.upper(), "RESET")}{text}{Fore.RESET}'


def check_installer_version():
    version = "v3.2"
    l_version_page = requests.get("https://github.com/CIO61/SHCE_Bootstrap_Installer/releases/latest")
    l_version = l_version_page.url.rpartition("/")[2]
    if version != l_version:
        try:
            download_file_w_progressbar("https://github.com/CIO61/SHCE_Bootstrap_Installer/releases/latest/download/update.zip",
                                        "update.zip", "Downloading Update")
            if os.path.exists("update.zip"):
                with open("update.zip", "rb") as updatefile:
                    if len(updatefile.read()) > 0:
                        os.startfile("selfupdate.exe")
                        sys.exit()
            pid = os.getpid()
            print(pid)
        except requests.RequestException:
            pass
        pass  # update available


def download_file_w_progressbar(url, filename, title):
    with open(filename, "wb") as f:
        try:
            response = requests.get(url, stream=True)
            if response.status_code == 404:
                return
            print(title)
            total_length = response.headers.get('content-length')
            if total_length is None:  # no content length header
                f.write(response.content)
            else:
                dl = 0
                total_length = int(total_length)
                for data in response.iter_content(chunk_size=4096):
                    dl += len(data)
                    f.write(data)
                    percentage = dl / total_length
                    done = int(50 * percentage)
                    forecolor = Fore.YELLOW if done != 50 else Fore.GREEN
                    sys.stdout.write(
                        f"\r[{forecolor}{'=' * done}{' ' * (50 - done)}{Fore.RESET}] [{percentage * 100:.2f}%]")
                    sys.stdout.flush()
                print()
        except requests.RequestException:
            with open("errors.log", "a") as errorfile:
                error_text = traceback.format_exc()
                pymsgbox.alert("Error Occured while downloading. A text file will now open containing details "
                               "about the error. Please share the contents of the file with the developer.")
                print(error_text, file=errorfile)
            sp.run(f"notepad errors.log")
            sys.exit()


def get_portable_git():
    portable_git_filename = "PortableGit-2.33.0.2-32-bit.7z.exe"
    url = f"https://github.com/git-for-windows/git/releases/download/v2.33.0.windows.2/{portable_git_filename}"
    download_file_w_progressbar(url, portable_git_filename, "Downloading Portable Git")
    sp.run(f"PortableGit-2.33.0.2-32-bit.7z.exe -y -o \"PortableGit\"")


def download_update(preview=False):
    repo_addr = "https://github.com/Krarilotus/BootstrapMultiplayerSetup.git"
    if not os.path.exists(f"{game_path}\\BootstrapMultiplayerSetup"):
        print("Downloading Bootstrap setup [First Time].")
        sp.run(f"{git_path} clone {repo_addr}", cwd=game_path)
    else:
        if os.path.exists(uninsjson := f"{game_path}\\BootstrapMultiplayerSetup\\uninstall.json"):
            os.remove(uninsjson)
        pull_update_kw = {"cwd": f"{game_path}\\BootstrapMultiplayerSetup", "creationflags": sp.CREATE_NO_WINDOW}
        sp.run(f"{git_path} reset --hard", **pull_update_kw)
        print("Checking for setup updates", end="")
        sys.stdout.flush()
        if preview:
            print("[Preview Version]")
            if "preview" not in sp.run(f"{git_path} remote", capture_output=True, text=True, **pull_update_kw).stdout.splitlines():
                sp.run(f"{git_path} remote add preview https://github.com/Altaruss28/BootstrapMultiplayerSetup.git", **pull_update_kw)
                sp.run(f"{git_path} checkout -b preview", **pull_update_kw)
            else:
                sp.run(f"{git_path} checkout preview", **pull_update_kw)
            sp.run(f"{git_path} pull preview main", **pull_update_kw)
            print("Switched to preview version")
        else:
            sp.run(f"{git_path} checkout main", **pull_update_kw)
            updcheck = sp.run(f"{git_path} pull", capture_output=True, text=True, **pull_update_kw)
            if updcheck.stdout.strip() == "Already up to date.":
                print("\nAlready on the latest version.")
            else:
                print("\nDone Checking Updates.")


def install_mod():
    print(f"Installing Bootstrap mode to {game_path}")
    with open(f"{game_path}\\BootstrapMultiplayerSetup\\version.txt") as versionfile:
        version = versionfile.read()
    print(f"Version: {version}")

    repo_path = f"{game_path}\\BootstrapMultiplayerSetup"

    p = sp.run(f'"{repo_path}\\installSetup.bat"', cwd=repo_path, creationflags=sp.CREATE_NEW_CONSOLE)

    if p.returncode:
        print("Installation Failed!")
        pymsgbox.alert("Installation Failed!")
        sys.exit(1)


def get_maps():
    files_in_dir = sp.run(f'dir /b /A-D "{game_path}\\mapsExtreme"', shell=True, text=True,
                          stdout=sp.PIPE).stdout.splitlines()
    files_in_dir = set(files_in_dir)
    matching_maps = files_in_dir & original_maps_set
    other_maps = files_in_dir - original_maps_set

    maps_to_copy = set(os.listdir(f"{game_path}\\BootstrapMultiplayerSetup\\DONTOPEN\\mapsExtreme"))
    maps_to_backup = other_maps - maps_to_copy
    for mapfile_original in matching_maps:
        pth_original_maps = f"{game_path}\\mapsExtreme\\original"
        if not os.path.exists(pth_original_maps):
            os.makedirs(pth_original_maps)
        shutil.move(f"{game_path}\\mapsExtreme\\{mapfile_original}", pth_original_maps)
    for mapfile_custom in maps_to_backup:
        pth_custom_maps = f"{game_path}\\mapsExtreme\\backup"
        if not os.path.exists(pth_custom_maps):
            os.makedirs(pth_custom_maps)
        try:
            shutil.move(f"{game_path}\\mapsExtreme\\{mapfile_custom}", pth_custom_maps)
        except shutil.Error:
            pass
    for map_to_copy in maps_to_copy:
        shutil.copy2(f"{game_path}\\BootstrapMultiplayerSetup\\DONTOPEN\\mapsExtreme\\{map_to_copy}",
                     f"{game_path}\\mapsExtreme")


def get_custom_graphics():
    if os.path.exists(cg_zip := f"{game_path}\\BootstrapMultiplayerSetup\\CustomGraphics.zip"):
        with zipfile.ZipFile(cg_zip) as zip_ref:
            zip_ref.extractall()
            for cg_folder in os.listdir("CustomGraphics"):
                cg_filename = "LocallyStoredTexture.gm1"
                with open(f"{game_path}\\gm\\{cg_folder}.gm1", 'rb') as gm1file_reference:
                    cg_file_checksum = hashlib.md5(gm1file_reference.read()).hexdigest()
                    for cg_file in os.listdir(f"CustomGraphics\\{cg_folder}"):
                        with open(f"CustomGraphics\\{cg_folder}\\{cg_file}", "rb") as cg_file_check:
                            if hashlib.md5(cg_file_check.read()).hexdigest() == cg_file_checksum:
                                break  # it is one of the known custom texture files
                    else:
                        i = 0
                        while cg_filename in os.listdir(f"CustomGraphics\\{cg_folder}"):
                            cg_filename = f"LocallyStoredTexture_{i}.gm1"
                        sp.run(f"copy {game_path}\\gm\\{cg_folder}.gm1 "
                               f"CustomGraphics\\{cg_folder}\\{cg_filename} > NUL", shell=True)


def conclude():
    shell = client.Dispatch("WScript.Shell")

    updater_shortcut_path = os.path.abspath(game_path + "\\BootstrapMod - Updater.lnk")
    preview_shortcut_path = os.path.abspath(game_path + "\\BootstrapMod - Preview Version.lnk")
    cg_cfg_shortcut_path = os.path.abspath(game_path + "\\BootstrapMod - CustomGraphicConfigurator.lnk")

    if not os.path.exists(updater_shortcut_path):
        shortcut_updater = shell.CreateShortCut(updater_shortcut_path)
        shortcut_updater.WorkingDirectory = os.path.abspath("")
        shortcut_updater.Targetpath = sys.executable
        shortcut_updater.WindowStyle = 1  # 7 - Minimized, 3 - Maximized, 1 - Normal
        shortcut_updater.save()

    if not os.path.exists(preview_shortcut_path):
        shortcut_updater = shell.CreateShortCut(preview_shortcut_path)
        shortcut_updater.WorkingDirectory = os.path.abspath("")
        shortcut_updater.Arguments = "preview"
        shortcut_updater.Targetpath = sys.executable
        shortcut_updater.WindowStyle = 1  # 7 - Minimized, 3 - Maximized, 1 - Normal
        shortcut_updater.save()

    cg_cfg_shortcut_check = os.path.exists(cg_cfg_shortcut_path)
    if not cg_cfg_shortcut_check:
        shortcut_cg_cfg = shell.CreateShortCut(cg_cfg_shortcut_path)
        shortcut_cg_cfg.WorkingDirectory = os.path.abspath(game_path)
        shortcut_cg_cfg.Targetpath = os.path.abspath("configure_custom_graphics.exe")
        shortcut_cg_cfg.WindowStyle = 1  # 7 - Minimized, 3 - Maximized, 1 - Normal
        shortcut_cg_cfg.save()

    print("Installation Completed!")
    pymsgbox.alert("Installation Completed" +
                   "\n\nIt seems you are using this installer version for the first time. \nYou can now use the "
                   "Updater and Custom Graphic Configurator in the game directory. You can move the Installer "
                   "anywhere you like. "
                   * (not cg_cfg_shortcut_check), "BootstrapMod - Installation Completed")


if __name__ == '__main__':
    preview_mode = "preview" in sys.argv
    colorama.init()
    check_installer_version()
    game_found = ("Stronghold_Crusader_Extreme.exe" in os.listdir(game_path))
    if not game_found:
        print(fc("red", "Bootstrap Mod is intended to be played with Stronghold Crusader Extreme. "
              "Put the installer in a game folder with Stronghold_Crusader_Extreme.exe."))
        pymsgbox.alert("Bootstrap Mod is intended to be played with Stronghold Crusader Extreme. "
                       "Put the installer in a game folder with Stronghold_Crusader_Extreme.exe.", "BootstrapMod")
        sys.exit()

    git_checker = sp.run("where git", stdout=sp.PIPE, stderr=sp.STDOUT)
    if git_checker.returncode == 0:
        git_path = "git"
    else:
        if not os.path.exists("PortableGit"):
            get_portable_git()
        git_path = f"PortableGit\\cmd\\git.exe"

    download_update(preview_mode)
    install_mod()
    get_maps()
    get_custom_graphics()
    conclude()
