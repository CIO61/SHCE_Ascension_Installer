import filecmp
import subprocess as sp
from collections import defaultdict

import sys

import colorama
from colorama.ansi import Fore, Cursor
import msvcrt
import os

colorama.init()
offset = 0  # scroll point of options menu

cgf = "AscensionMod\\Graphics"
cgf_src = "Ascension\\Graphics"
backupdir = f"{cgf}\\local"
os.makedirs(cgf, exist_ok=True)

if "Stronghold_Crusader_Extreme.exe" not in os.listdir():
    sys.exit()

option_list = defaultdict(list)
sp.run(f"robocopy /E {cgf_src} {cgf}", stdout=sp.DEVNULL, stderr=sp.DEVNULL)
for folder in os.listdir(cgf):
    for file in [_f.removesuffix(".gm1") for _f in os.listdir(f"{cgf}\\{folder}")]:
        option_list[file].append(folder)


for key, val in list(option_list.items()):
    if len(val) < 2:
        option_list.pop(key)


if not os.path.exists(f"{backupdir}"):
    sp.run(f"robocopy gm {backupdir}", stdout=sp.DEVNULL, stderr=sp.DEVNULL)


selection = {}


def check_selection():
    for key, val in option_list.items():
        keepcheck = True
        for j, item in enumerate(val):
            same = keepcheck and filecmp.cmp(f"{cgf}\\{item}\\{key}.gm1", f"gm\\{key}.gm1")
            if same:
                keepcheck = False
                selection[key] = j
        if keepcheck:
            selection[key] = -1


def prepare_text():
    ret_text = ""
    for i, kvpair in enumerate(option_list.items()):
        _k, _v = kvpair
        item_txts = []
        for j, item in enumerate(_v):
            color = Fore.GREEN if selection[_k] == j else Fore.RED
            item_txts.append(f"{color}{item}{Fore.RESET}")
        number = f"{Fore.CYAN}{i-offset+1}{Fore.RESET}" if (0 < i-offset+1 <= 9) else " "
        ret_text += f"[{number}] {_k}: [{' '.join(item_txts)}]\n"
    return ret_text


def print_status(rewind=False):
    if rewind:
        print()
        print(Cursor.UP(len(selection)+4), end="")
    print(f"Custom Texture Settings ({Fore.GREEN}green{Fore.RESET} shows the currently active option)")
    print(prepare_text())
    print(f"[{Fore.CYAN}0{Fore.RESET}]: Exit")


def input_loop():
    global offset
    skip = False
    ch = ""
    while ch.upper() not in ["0"]:
        ch = msvcrt.getwch()
        if ch == "à":
            skip = True
            continue
        if skip:
            skip = False
            if ch == "P" and offset < len(selection)-9:
                offset += 1
            elif ch == "H" and offset >= 1:
                offset -= 1
            else:
                ch = ""
                continue

        sel_keys = list(selection.keys())
        if ch in [str(n) for n in range(1, len(sel_keys)+1)]:
            category = sel_keys[int(ch)-1+offset]
            selection[category] += 1
            selection[category] %= len(option_list[category])
            sel_idx = selection[category]
            file = option_list[category][sel_idx]
            src_file = f"{cgf}\\{file}\\{category}.gm1"
            tgt_file = f"gm\\{category}.gm1"
            same = filecmp.cmp(src_file, tgt_file)
            sp.run(f'copy /Y "{src_file}" "{tgt_file}" > NUL', shell=True)

        print_status(rewind=True)


print(f"Welcome to custom graphic configurator.\n"
      f"Use {Fore.CYAN}number keys{Fore.RESET} on your keyboard to apply the options.\n"
      f"Use {Fore.MAGENTA}arrow keys{Fore.RESET} to scroll within the menu")
check_selection()
print_status()
input_loop()
