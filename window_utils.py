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
SC_MINIMIZE = 0xF020
WM_SYSCOMMAND = 0x0112
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOPMOST = 0x00000008
GWL_STYLE = -16
GW_OWNER = 4
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000


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


def get_all_user_windows():
    """
    Get windows that are visible in the taskbar (i.e. user-accessible).
    """
    windows = {}

    def enum_windows_proc_callback(hwnd, l_param):
        if not ctypes.windll.user32.IsWindowVisible(hwnd):
            return True

        style = ctypes.windll.user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        if style & WS_EX_TOOLWINDOW:
            return True

        if ctypes.windll.user32.GetWindow(hwnd, GW_OWNER):
            return True

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

    :param hwnd: The window handle (hwnd) of the window.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow
    """
    ctypes.windll.user32.ShowWindow(hwnd, SW_MINIMIZE)


def minimize_window_syscommand(hwnd: int):
    """
    Minimize a window using WM_SYSCOMMAND message (and a window handle). \n
    The previous `minimize_window()` function exhibited some strange behavior:
    - The user's focus would be lost upon minimizing the window.
    - The desktop would show milliseconds of darkness after minimizing the window.
    

    :param hwnd: The window handle (hwnd) of the window.

    Relevant win32 documentation:
    ----------------------------
    https://learn.microsoft.com/en-us/windows/win32/menurc/wm-syscommand
    https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-sendmessage
    """
    ctypes.windll.user32.SendMessageW(hwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0)


def get_process_name_from_hwnd(hwnd: int) -> str:
    """
    Get the name of the process associated with a window handle (hwnd) without using psutil.

    :param hwnd: The window handle (hwnd) of the window.
    :return: The process name or an empty string if not found.
    """
    pid = ctypes.c_ulong()
    ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

    kernel32 = ctypes.windll.kernel32
    process_handle = kernel32.OpenProcess(
        PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value
    )

    if not process_handle:
        return ""

    buffer = ctypes.create_unicode_buffer(260)
    size = ctypes.c_ulong(260)
    if kernel32.QueryFullProcessImageNameW(
        process_handle, 0, buffer, ctypes.byref(size)
    ):
        full_path = buffer.value
        process_name = full_path.split("\\")[-1]
    else:
        process_name = ""
    kernel32.CloseHandle(process_handle)
    return process_name
