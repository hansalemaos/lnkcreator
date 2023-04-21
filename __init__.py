import argparse
import ast
import sys
import tempfile
import os
import pathlib
import subprocess
from typing import Union

from subprocesshidden import Popen


def escapepa(p):
    return os.path.normpath(p).replace("\\", "\\\\")


def escapeargu(p):
    if not p:
        return []
    if not isinstance(p, list):
        p = [p]
    allali = []
    for pp in p:
        allali.append(pp.replace("\\", "\\\\").replace('"', '\\"'))
    return allali


def create_shortcut(
    shortcut_path: str,
    target: str,
    arguments: list,
    hotkey="",
    working_dir: Union[str, None] = None,
    minimized_maximized_normal: str = "minimized",
    asadmin: bool = False,
) -> str:
    """
    Creates a Windows shortcut (.lnk) file at the specified path with the specified properties.

    Args:
        shortcut_path (str): The path where the shortcut file will be created.
        target (str): The path of the target file or application that the shortcut will point to.
        arguments (list): A list of arguments to be passed to the target file or application.
        hotkey (str, optional): The hotkey combination to activate the shortcut. Defaults to "".
        working_dir (Union[str, None], optional): The working directory for the target file or application. Defaults to None.
        minimized_maximized_normal (str, optional): The window state of the target application when the shortcut is activated.
            Possible values are "minimized", "maximized", or "normal". Defaults to "minimized".
        asadmin (bool, optional): If True, the shortcut will be created with administrative privileges. Defaults to False.

    Returns:
        str: The JavaScript content used to create the shortcut.

    Raises:
        OSError: If the shortcut file cannot be created.


    """

    wco = 4

    if minimized_maximized_normal == "minimized":
        wco = 7
    elif minimized_maximized_normal == "maximized":
        wco = 0

    shortcut_path2 = pathlib.Path(shortcut_path)
    shortcut_path2.parent.mkdir(parents=True, exist_ok=True)
    if str(working_dir) == "None":
        working_dir = None
    if not working_dir:
        working_dir = os.path.normpath(os.path.dirname(target))
    working_dir = escapepa(working_dir)
    shortcut_path = escapepa(shortcut_path)
    target = escapepa(target)

    if not arguments:
        arguments = ""
    if isinstance(arguments, list):
        arguments = " ".join(arguments)

    if isinstance(arguments, list):
        arguments = " ".join(arguments)

    arguments = " ".join(escapeargu(arguments))
    js_content = f"""
        var sh = WScript.CreateObject("WScript.Shell");
        var shortcut = sh.CreateShortcut("{shortcut_path}");
        shortcut.WindowStyle = {wco};
        shortcut.TargetPath = "{target}";
        shortcut.Hotkey = "{hotkey}";
        shortcut.Arguments = "{arguments}";
        shortcut.WorkingDirectory = "{working_dir}";
        shortcut.IconLocation = "{target}";
        shortcut.Save();"""
    fd, path = tempfile.mkstemp(".js")
    try:
        with os.fdopen(fd, "w") as f:
            f.write(js_content)

        p = Popen(
            [r"wscript.exe", path],
            shell=False,
            stderr=subprocess.PIPE,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
        )
        p.wait(5)

    finally:
        try:
            os.unlink(path)
        except Exception as fe:
            pass

    if asadmin:
        with open(shortcut_path, "rb") as f:
            ba = bytearray(f.read())
        ba[0x15] = ba[0x15] | 0x20
        with open(shortcut_path, "wb") as f:
            f.write(ba)
    return js_content


def asteval(x):
    try:
        return ast.literal_eval(str(x))
    except Exception:
        return x


def npath(x):
    try:
        return os.path.normpath(str(x))
    except Exception:
        return x


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        formatter_class=lambda prog: argparse.HelpFormatter(prog, max_help_position=65)
    )

    parser.add_argument(
        "--shortcut_path",
        help=r"""The path where the shortcut file will be created.""",
    )
    parser.add_argument(
        "--target",
        default="",
        help=r"""The path to the target file or application that the shortcut will launch.""",
    )

    parser.add_argument(
        "--hotkey",
        default="",
        help=r"""Hotkey for opening the lnk file. Defaults to ''.""",
    )
    parser.add_argument(
        "--working_dir",
        default="None",
        help=r"""The working directory for the target file or application.  (Working dict of shortcut_path).""",
    )
    parser.add_argument(
        "--mode",
        default="minimized",
        help=r"""The window state of the target application when launched. Can be "normal", "minimized" or "maximized" """,
    )
    parser.add_argument(
        "--asadmin",
        default="False",
        help=r"""Whether to run the shortcut as an administrator. Defaults to False.""",
    )
    parser.add_argument(
        r"--args",
        default="",
        help=r"""Arguments to be passed to the target file or application as they would receive them. Example: 
        lnkgonewild.exe --shortcut_path C:\Users\hansc\Desktop\testlink.lnk --target C:\cygwin\bin\lsattr.exe --hotkey Ctrl+Alt+q --working_dir None --mode minimized --silentlog None --asadmin False --arguments -a -d""",
    )

    if len(sys.argv) < 3:
        parser.print_help()
    else:
        try:
            argsindex = sys.argv.index("--args")
            args_ = sys.argv[argsindex + 1 :].copy()
            sys.argv = sys.argv[:argsindex]
            args_ = " ".join(args_).strip()
            print(args_)
        except Exception:
            args_ = "[]"

        args = parser.parse_args()
        shortcut_path = npath(asteval(args.shortcut_path))
        target = npath(asteval(args.target))
        arguments = asteval(args_)
        hotkey = asteval(args.hotkey)
        working_dir = npath(asteval(args.working_dir))
        minimized_maximized_normal_invisible = asteval(args.mode)
        asadmin = asteval(args.asadmin)
        create_shortcut(
            shortcut_path=shortcut_path,
            target=target,
            arguments=arguments,
            minimized_maximized_normal=minimized_maximized_normal_invisible,  #
            asadmin=asadmin,  # enables the admin check box
            hotkey=hotkey,
            working_dir=working_dir,
        )
