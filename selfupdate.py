import os
import shutil

if os.path.exists("update.zip"):
    print("Updating Installer")
    shutil.unpack_archive("update.zip")
    os.remove("update.zip")
    os.startfile("BootstrapInstaller.exe")
