print("Cyber-Guide-Dog 赛博导盲犬 v1.0.0")
import os, winshell, conkits, colorama, msvcrt
import zipfile, rarfile, py7zr
from win32com.client import Dispatch
from rich import print

_print = print


def relative_file(path):
    return os.path.join(os.path.dirname(__file__), path)


def use_game_database(*path):
    return os.path.join(game_database, *path)


def print(*strs):
    _print(f"[white bold]{''.join(strs)}[/white bold]")


colorama.init()
print("")
print("Searching for game archive...")
game_database = os.path.expanduser("~/.cyber-guide-dog")
os.makedirs(game_database, exist_ok=True)
rarfile.UNRAR_TOOL = relative_file("unrar.exe")
archive_list = []
for root, dirs, files in os.walk("."):
    for i in files:
        file = None
        try:
            file = zipfile.ZipFile(i)
        except:
            try:
                file = rarfile.RarFile(i)
            except:
                try:
                    file = py7zr.SevenZipFile(i)
                except:
                    file = None
        if file:
            archive_list.append(file)
for i in archive_list:
    i: zipfile.ZipFile
    i_ext = os.path.splitext(i.filename)
    game_path = use_game_database(i_ext[0])
    print(f"Found archive [green][{i.filename}][/green], extracting...")
    i.extractall(game_path)
    print("Finding game entry...")
    game_main = []
    for root, dirs, files in os.walk(game_path):
        for file in files:
            if os.path.splitext(file)[1].upper() == ".EXE":
                game_main.append(file)
    if len(game_main)==0:
        print(f"Cannot find game entry, skipping {i.filename}.")
        continue
    game_main.insert(0,"Skip")
    selector = conkits.Choice(options=game_main)
    print(
        "Found lots of game entry, select which one is right(If you can't read HUMAN'S TEXT, ask others)."
    )
    print("[yellow](Use arrow keys ↑↓)[/yellow]")
    selected=selector.run()
    if selected==0:
        print("[red]Skipped.[/red]")
    else:
        game_main = game_main[selected]
        print(
            f"Selected [green]\\[{os.path.splitext(game_main)[0]}][/green], creating desktop shortcut..."
        )
        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(os.path.join(winshell.desktop(), i_ext[0] + ".lnk"))
        shortcut.Targetpath = use_game_database(i_ext[0], game_main)
        shortcut.WorkingDirectory = game_path
        shortcut.save()
print("Done!")
print("Press any key to exit...")
msvcrt.getch()
