import os
import subprocess as sp
import sys
import traceback
import threading
import time
import json
from urllib import request
from urllib.error import URLError
import pymsgbox
import win32con
import shutil
import zipfile
import hashlib

# noinspection PyPackageRequirements
from win32gui import GetOpenFileNameW
# noinspection PyPackageRequirements
from win32com import client

configfile_path = f"installer_config.json"
original_maps_set = {
    "Best_Friends.map", "Bird_in_flight.map", "Coastal_Trap.map", "Crossroads.map", "Divided.map", "Enclosure.map",
    "Fury_Bay.map", "Jealous_Neighbours.map", "MP-No mans land.map", "MP-No where to hide.map", "MP-Slopes of doom.map",
    "Lionheart.map", "Look_out.map", "MP-Divided.map", "MP-Downhill Scrum.map", "MP-Snake River.map",
    "MP-Surrounded.map", "MP-The rocky divide.map", "MP-Two Falls.map", "MP-Valley of the lords.map",
    "Phoenix.map", "Rivers_Fork.map", "Snake_River.map", "Spider_Island.map", "Swampy_Island.map",
    "The_Host.map", "Three_Little_Pigs.map", "Ultimate_Victory.map", "Wazirs_Fortress.map", "Wide_Open_Plain.map"
}
original_gm1_md5 = {
    'body_arab_assasin': '5f988b6df71d8b30a975f93ed4e758a3',
    'body_horse_archer_top': '9cf44ef27eb6c9b15cdcc1a1edbf54f6',
    'body_fighting_monk': 'a9968074978a3103f8343f2456d87b91',
    'tree_shrub2': '7281ddf03f7307a6d1fba0eef785090e'}


def check_for_maps_extreme_content(path):
    files_in_dir = sp.run(f'dir /b /A-D "{path}\\mapsExtreme"', shell=True, text=True, stdout=sp.PIPE).stdout.splitlines()
    files_in_dir = set(files_in_dir)
    matching_maps = files_in_dir & original_maps_set
    other_maps = files_in_dir - original_maps_set
    return matching_maps, other_maps


def get_portable_git():
    portable_git_filename = "PortableGit-2.33.0.2-32-bit.7z.exe"
    if not os.path.exists(f"{working_directory}\\{portable_git_filename}"):
        try:
            url = "https://github.com/git-for-windows/git/releases/download/v2.33.0.windows.2/PortableGit-2.33.0.2-32-bit.7z.exe"
            request.urlretrieve(url, f"{working_directory}\\{portable_git_filename}")
        except URLError:
            with open(f"{working_directory}\\errors.log", "a") as errorfile:
                error_text = traceback.format_exc()
                pymsgbox.alert("Some Error Occured while setting up. A text file will now open containing details "
                               "about the error. Please share the contents of the file with the developer.")

                print(error_text, file=errorfile)
            sp.run(f"notepad {working_directory}\\errors.log")
            sys.exit()
    sp.run(f"{working_directory}\\PortableGit-2.33.0.2-32-bit.7z.exe -y -o \"PortableGit\"", cwd=working_directory)


def get_game_path():
    try:
        fname, customfilter, flags = GetOpenFileNameW(
            InitialDir=os.path.abspath(os.curdir),
            Flags=win32con.OFN_EXPLORER | win32con.OFN_FILEMUSTEXIST,
            File='Stronghold_Crusader_Extreme.exe', DefExt='exe',
            Title='Locate Stronghold_Crusader_Extreme.exe',
            Filter='Stronghold Crusader Extreme\0Stronghold_Crusader_Extreme.exe',
            FilterIndex=0)
        if not fname.endswith("Stronghold_Crusader_Extreme.exe"):
            pymsgbox.alert("Stronghold_Crusader_Extreme.exe was not shown. Aborting install.")
            sys.exit(0)
    except Exception:
        pymsgbox.alert("Stronghold_Crusader_Extreme.exe was not shown. Aborting install.")
        sys.exit(0)
    return fname.rpartition("\\")[0]


# STEP 0: CHECK EXISTING CONFIG
scriptpath = (" ".join(sys.argv[1:]) if len(sys.argv) > 1 else "")
if os.path.exists(configfile_path):
    with open(configfile_path) as cfg_file:
        cfg_data = json.load(cfg_file)
    game_path = cfg_data["game_path"]
    working_directory = cfg_data["working_directory"]
    no_prompt = cfg_data["no_prompt"]
else:
    game_path = scriptpath
    working_directory = ""
    no_prompt = False


# STEP 1: FIND OUT THE GAME PATH

# 1.1: In the same folder as the installer
if os.path.abspath(scriptpath) == os.path.abspath(game_path):
    if os.path.isdir(scriptpath) and ("Stronghold_Crusader_Extreme.exe" in os.listdir(scriptpath)):
        if not no_prompt:
            rv = pymsgbox.confirm(text=f"Found SHCE here for installation. Proceed?", title="Bootstrap Installer")
            if rv == "OK":
                no_prompt = True
                game_path = scriptpath

# 1.2: In the folder specified by config file
game_found = os.path.isdir(game_path) and ("Stronghold_Crusader_Extreme.exe" in os.listdir(game_path))
if game_found:
    if not no_prompt:
        rv = pymsgbox.confirm(text=f"Found SHCE in {game_path} for installation from previous config. Proceed?", title="Bootstrap Installler")
        if rv == "OK":
            rv2 = pymsgbox.confirm(text=f"Do you want to always use this path?", title="Bootstrap Installer",
                                   buttons=[pymsgbox.YES_TEXT, pymsgbox.NO_TEXT])
            no_prompt = (rv2 == "Yes")
        else:
            game_path = get_game_path()
else:
    pymsgbox.alert("Cannot find game installation! Please locate Stronghold_Crusader_Extreme.exe")
    game_path = get_game_path()

# STEP 2: SET WORKING DIRECTORY
if not working_directory:
    if game_path[0] != os.environ["USERPROFILE"][0]:
        working_directory = f"{game_path[0]}:\\bootstrap_installer_workspace"
    else:
        working_directory = f'{os.environ["USERPROFILE"]}\\bootstrap_installer_workspace'

os.makedirs(working_directory, exist_ok=True)
print(f"Setting working directory to: [{working_directory}]")

# STEP 3: GET PORTABLE GIT
git_checker = sp.run("where git", stdout=sp.PIPE, stderr=sp.STDOUT, text=True)
if git_checker.returncode == 0:
    git_path = "git"
else:
    if not os.path.exists(f"{working_directory}\\PortableGit"):
        print("Performing First Time Setup... This will take a while, please wait!")
        portable_git_getter = threading.Thread(target=get_portable_git)
        portable_git_getter.start()

        print("Preparing update system.")
        time.sleep(2)
        while portable_git_getter.is_alive():
            time.sleep(1)
            print(".", end="")
        else:
            print("\n Done!")
    git_path = f"{working_directory}\\PortableGit\\cmd\\git.exe"

# STEP 4: GET BOOTSTRAP MP SETUP
if not os.path.exists(f"{working_directory}\\BootstrapMultiplayerSetup"):
    repo_addr = "https://github.com/Krarilotus/BootstrapMultiplayerSetup.git"
    git_cloner = threading.Thread(target=sp.run, kwargs={
        "args": f"{git_path} clone {repo_addr}",
        "cwd": working_directory
    })
    git_cloner.start()

    print("Downloading Bootstrap setup [First Time].")
    time.sleep(2)
    while git_cloner.is_alive():
        time.sleep(1)
        print(".", end="")
    else:
        print("\n Done!")
else:
    # STEP 4.5: CHECKOUT&PULL IN CASE IT HAS BEEN GOTTEN BEFORE BUT NOT UP TO DATE/CORRUPTED
    if os.path.exists(uninsjson := f"{working_directory}\\BootstrapMultiplayerSetup\\uninstall.json"):
        os.remove(uninsjson)
    git_checkouter = threading.Thread(target=sp.run, kwargs={
        "args": f"{git_path} checkout main .",
        "cwd": f"{working_directory}\\BootstrapMultiplayerSetup",
        "creationflags": sp.CREATE_NO_WINDOW
    })
    git_checkouter.start()
    print("Checking for setup updates", end="")
    time.sleep(2)
    while git_checkouter.is_alive():
        time.sleep(1)
        print(".", end="")
    else:
        print("\nDone Checking Updates.")

    git_puller = threading.Thread(target=sp.run, kwargs={
        "args": f"{git_path} pull",
        "cwd": f"{working_directory}\\BootstrapMultiplayerSetup",
        "creationflags": sp.CREATE_NO_WINDOW
    })
    git_puller.start()
    while git_puller.is_alive():
        time.sleep(1)

# STEP 5: CREATE A DIRECTORY JUNCTION TO GAMEDIR <<==>> BootstrapMultiplayerSetup
print(f"Installing Bootstrap mode to {game_path}")
with open(f"{working_directory}\\BootstrapMultiplayerSetup\\version.txt") as versionfile:
    version = versionfile.read()
print(f"Version: {version}")

foldername = "BootstrapInstallerFiles"
temp_path = f"{game_path}\\{foldername}"
source_path = f"{working_directory}\\BootstrapMultiplayerSetup"

# STEP 6: RUN INSTALLSETUP.BAT, from junction dir
if os.path.exists(f"{temp_path}"):
    sp.run(f"rmdir {foldername}", cwd=game_path, shell=True)
sp.run(f'cmd /c mklink /J "{temp_path}" "{source_path}"', creationflags=sp.CREATE_NO_WINDOW)
p = sp.run(f'"{temp_path}\\installSetup.bat"', cwd=temp_path, creationflags=sp.CREATE_NEW_CONSOLE)

if p.returncode:
    print("Installation Failed!")
    pymsgbox.alert("Installation Failed!")
    sys.exit(1)

# STEP 7: COPY OVER MAPS
matching, other = check_for_maps_extreme_content(game_path)
maps_to_copy = set(os.listdir(f"{working_directory}\\BootstrapMultiplayerSetup\\DONTOPEN\\mapsExtreme"))
maps_to_backup = other - maps_to_copy
for mapfile_original in matching:
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
    shutil.copy2(f"{working_directory}\\BootstrapMultiplayerSetup\\DONTOPEN\\mapsExtreme\\{map_to_copy}",
                 f"{game_path}\\mapsExtreme")

# STEP 8: ACQUIRE CUSTOM GRAPHICS
if os.path.exists(f"{working_directory}\\BootstrapMultiplayerSetup\\CustomGraphics.zip"):
    with zipfile.ZipFile(f"{working_directory}\\BootstrapMultiplayerSetup\\customGraphics.zip") as zip_ref:
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


# STEP 9: VOILA, save config
with open(configfile_path, "w") as configfile_write:
    json.dump({
        "game_path": game_path,
        "working_directory": working_directory,
        "no_prompt": no_prompt
    }, configfile_write)

shell = client.Dispatch("WScript.Shell")

updater_shortcut_path = os.path.abspath(game_path+"\\BootstrapMod - Updater.lnk")
cg_cfg_shortcut_path = os.path.abspath(game_path+"\\BootstrapMod - CustomGraphicConfigurator.lnk")

if not os.path.exists(updater_shortcut_path):
    shortcut_updater = shell.CreateShortCut(updater_shortcut_path)
    shortcut_updater.WorkingDirectory = os.path.abspath("")
    shortcut_updater.Arguments = os.path.abspath("..\\")
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
               "\n\nIt seems you are using this installer version for the first time. \nYou can now use the Updater "
               "and Custom Graphic Configurator in the game directory. You can move the Installer anywhere you like."
               * (not cg_cfg_shortcut_check), "BootstrapMod - Installation Completed")

os.startfile(game_path)
