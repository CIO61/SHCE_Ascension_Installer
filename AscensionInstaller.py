import os
import subprocess as sp
import sys
import traceback
from functools import partial

import colorama
import pymsgbox
import requests
import shutil

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


bad = partial(fc, "red")
neutral = partial(fc, "yellow")
good = partial(fc, "green")


def check_installer_version():
    version = "v5.0"
    source_address = "https://github.com/CIO61/SHCE_Ascension_Installer"
    l_version_page = requests.get(f"{source_address}/releases/latest")
    l_version = l_version_page.url.rpartition("/")[2]
    if version != l_version:
        try:
            download_file_w_progressbar(f"{source_address}/releases/latest/download/update_v2.zip",
                                        "update_v2.zip", "Downloading Update")
            if os.path.exists("update_v2.zip"):
                with open("update_v2.zip", "rb") as updatefile:
                    if len(updatefile.read()) > 0:
                        with open("leftover_command.txt", "w") as cmdfile:
                            cmdfile.write(" ".join(sys.argv[1:]))
                        sp.Popen('timeout 1 > NUL && selfupdate.exe', shell=True)
                        sys.exit()
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


def download_update():
    pull_upd = partial(sp.run, cwd=f"{game_path}\\{workfolder}", creationflags=sp.CREATE_NO_WINDOW)

    if not os.path.exists(f"{game_path}\\{workfolder}"):
        print("Downloading Ascension setup.")
        sp.run(f"{git_path} clone {repo_addr} {workfolder}", cwd=game_path)
    else:
        if os.path.exists(uninsjson := f"{game_path}\\{workfolder}\\uninstall.json"):
            os.remove(uninsjson)
        print("Checking for setup updates", end="")
        sys.stdout.flush()
        pull_upd(f"{git_path} reset --hard", stdout=sp.DEVNULL)
        pull_upd(f"{git_path} clean -fdx", stdout=sp.DEVNULL)
        updcheck = pull_upd(f"{git_path} pull", capture_output=True, text=True)
        if updcheck.stdout.strip() == "Already up to date.":
            print("\nAlready on the latest version.")
        else:
            print("\nDone Checking Updates.")


def install_mod():
    print(f"Installing UCP...")
    repo_path = f"{game_path}\\{workfolder}"

    # apply UCP
    ucp = sp.Popen(f'{repo_path}\\UCP\\UnofficialCrusaderPatchCLI.exe "--language=1" "--path={game_path}"',
                   cwd=f"{repo_path}\\UCP", text=True, stdin=sp.PIPE, stdout=sp.PIPE, stderr=sp.PIPE)

    line = ucp.stdout.readline()
    while line:
        if "ucp successfully installed" in line.lower():
            print(good("UCP successfully installed"))
            break
        if "Custom AIVs detected." in line:
            line, err = ucp.communicate("delete\n")
        else:
            line = ucp.stdout.readline()
    ucp.wait()

    if ucp.returncode:
        print(bad("UCP Installation Failed!"))
        pymsgbox.alert("UCP Installation Failed!")
        sys.exit(1)

    # apply mod patch
    print(f"Installing balance patch...")
    with open(f"{game_path}\\{workfolder}\\Documentation\\version.txt") as versionfile:
        version = versionfile.read().strip()
    print(neutral(f"Version: {version}"))

    mod = sp.run(f'{repo_path}\\Rebalance\\mod.exe -gamepath {game_path}', cwd=f'{repo_path}\\Rebalance\\')

    if mod.returncode:
        print(bad("Balance Patch Failed!"))
        pymsgbox.alert("Balance Patch Failed!")
        sys.exit(1)


def get_maps():
    files_in_dir = sp.run(f'dir /b /A-D "{game_path}\\mapsExtreme"', shell=True, text=True,
                          stdout=sp.PIPE).stdout.splitlines()
    files_in_dir = set(files_in_dir)
    matching_maps = files_in_dir & original_maps_set
    other_maps = files_in_dir - original_maps_set

    maps_to_copy = set(os.listdir(f"{game_path}\\{workfolder}\\Maps"))
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
        sp.run(f'move /Y "{game_path}\\mapsExtreme\\{mapfile_custom}" "{pth_custom_maps}"',
               shell=True, stdout=sp.DEVNULL)
    for map_to_copy in maps_to_copy:
        shutil.copy2(f"{game_path}\\{workfolder}\\Maps\\{map_to_copy}", f"{game_path}\\mapsExtreme")


def apply_localisation():
    pass


def conclude():
    shell = client.Dispatch("WScript.Shell")

    updater_shortcut_path = os.path.abspath(game_path + "\\AscensionMod - Updater.lnk")
    cg_cfg_shortcut_path = os.path.abspath(game_path + "\\AscensionMod - CustomGraphicConfigurator.lnk")

    def create_shortcut(shortcut_path, *,
                        working_dir=os.path.abspath(""),
                        arguments="",
                        target_path=sys.executable,
                        icon_path=os.path.abspath("icon.ico")):
        shortcut_file = shell.CreateShortCut(shortcut_path)
        shortcut_file.WorkingDirectory = working_dir
        shortcut_file.Arguments = arguments
        shortcut_file.Targetpath = target_path
        shortcut_file.IconLocation = icon_path
        shortcut_file.WindowStyle = 1  # 7 - Minimized, 3 - Maximized, 1 - Normal
        shortcut_file.save()

    if not os.path.exists(updater_shortcut_path):
        create_shortcut(updater_shortcut_path)

    cg_cfg_shortcut_check = os.path.exists(cg_cfg_shortcut_path)
    if not cg_cfg_shortcut_check:
        create_shortcut(cg_cfg_shortcut_path,
                        working_dir=os.path.abspath(game_path),
                        target_path=os.path.abspath("configure_custom_graphics.exe"))

    print(good("Installation Completed!"))
    pymsgbox.alert("Installation Completed" +
                   "\n\nIt seems you are using this installer version for the first time. \nYou can now use the "
                   "Updater and Custom Graphic Configurator in the game directory. You can move the Installer "
                   "anywhere you like. "
                   * (not cg_cfg_shortcut_check), "AscensionMod - Installation Completed")


if __name__ == '__main__':
    repo_addr = "https://github.com/Krarilotus/Ascension.git"
    workfolder = "Ascension"

    # backup folder for graphics and localisation
    os.makedirs("backups", exist_ok=True)

    colorama.init()
    check_installer_version()
    game_found = ("Stronghold_Crusader_Extreme.exe" in os.listdir(game_path))
    if not game_found:
        print(bad("Ascension Mod is intended to be played with Stronghold Crusader Extreme. "
                  "Put the installer in a game folder with Stronghold_Crusader_Extreme.exe."))
        pymsgbox.alert("Ascension Mod is intended to be played with Stronghold Crusader Extreme. "
                       "Put the installer in a game folder with Stronghold_Crusader_Extreme.exe.", "Mod")
        sys.exit()

    git_checker = sp.run("where git", stdout=sp.PIPE, stderr=sp.STDOUT)
    if git_checker.returncode == 0:
        git_path = "git"
    else:
        if not os.path.exists("PortableGit"):
            get_portable_git()
        git_path = f"PortableGit\\cmd\\git.exe"

    download_update()
    install_mod()
    get_maps()
    apply_localisation()
    conclude()
