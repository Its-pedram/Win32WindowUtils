"""
A collection of utility functions for working with windows (opacity, focus, etc) in Python using ctypes.
"""

import ctypes

WS_EX_LAYERED = 0x00080000  # https://learn.microsoft.com/en-us/windows/win32/winmsg/extended-window-styles
LWA_ALPHA = 0x00000002  # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setlayeredwindowattributes
GWL_EXSTYLE = (
    -20
)  # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlongptra
SW_MINIMIZE = 6  # https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow

def get_hwnd_from_title(title: str) -> int:
    """
    Get the window handle (hwnd) of a window by its title.

    :param title: The title of the window.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-findwindoww
    """
    hwnd = ctypes.windll.user32.FindWindowW(None, title)
    return hwnd


def get_title_from_hwnd(hwnd: int) -> str:
    """
    Get the title of a window by its window handle (hwnd).

    :param hwnd: The window handle (hwnd) of the window.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtextlengthw \n
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowtextw
    """
    length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
    buff = ctypes.create_unicode_buffer(length + 1)
    ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
    return buff.value


def get_all_windows():
    """
    Get all the windows that are currently open.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumwindows \n
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-iswindowvisible
    """
    windows = {}

    def enum_windows_proc_callback(hwnd, l_param):
        if ctypes.windll.user32.IsWindowVisible(hwnd):
            title = get_title_from_hwnd(hwnd)
            if title:
                windows[hwnd] = title
        return True

    ctypes.windll.user32.EnumWindows(
        ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_int, ctypes.c_int)(
            enum_windows_proc_callback
        ),
        0,
    )
    return windows


def set_window_opacity(hwnd: int, opacity: int):
    """
    Set the opacity of a window.

    :param hwnd: The window handle (hwnd) of the window.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getwindowlongptra \n
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowlongptra \n
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setlayeredwindowattributes
    """
    # Make the window layered (adds WS_EX_LAYERED to the extended window style)
    ctypes.windll.user32.SetWindowLongPtrA(
        hwnd,
        GWL_EXSTYLE,
        ctypes.windll.user32.GetWindowLongPtrA(hwnd, GWL_EXSTYLE) | WS_EX_LAYERED,
    )
    ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, opacity, LWA_ALPHA)


def get_focused_window() -> int:
    """
    Get the window handle (hwnd) and title of the currently focused window.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getforegroundwindow
    """
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    return hwnd


def minimize_window(hwnd: int):
    """
    Minimize a window by its window handle (hwnd).

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
    """
    ctypes.windll.user32.ShowWindow(hwnd, SW_MINIMIZE)
