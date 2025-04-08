# Add lifespan support for startup/shutdown with strong typing
from contextlib import asynccontextmanager
from collections.abc import AsyncIterator
from dataclasses import dataclass
from typing import Literal, Any, TypedDict, Optional, Union

from ahk import AsyncAHK, AsyncWindow, TitleMatchMode
import wmutil
from mcp.server.fastmcp import Context, FastMCP
import easyocr
from mss import mss
from mss.base import MSSBase
import numpy as np


@dataclass
class AppContext:
    ahk: AsyncAHK


@asynccontextmanager
async def app_lifespan(server: FastMCP) -> AsyncIterator[AppContext]:
    """Manage application lifecycle with type-safe context"""
    # Initialize on startup
    from ahk._async.transport import AsyncDaemonProcessTransport
    ahk = AsyncAHK(version='v1')
    await ahk.get_mouse_position()
    try:
        yield AppContext(ahk=ahk)
    finally:
        # Cleanup on shutdown
        try:
            ahk._transport: AsyncDaemonProcessTransport
            if ahk._transport._proc is not None:
                ahk._transport._proc.kill()
            ahk.stop_hotkeys()
        except:
            pass


# Specify dependencies for deployment and development
mcp = FastMCP("AHK MCP", dependencies=["ahk", "ahk-binary", "mss", "easyocr", "numpy", "wmutil"], lifespan=app_lifespan)

class WindowInfo(TypedDict):
    pid: int
    title: str
    process_path: str
    process_name: str
    ahk_id: str
    x_position: int
    y_position: int
    height: int
    width: int

async def window_to_info(win: AsyncWindow) -> WindowInfo:
    pos = await win.get_position()
    win_info: WindowInfo = {
        'pid': await win.get_pid(),
        'title': await win.get_title(),
        'process_path': await win.get_process_path(),
        'process_name': await win.get_process_name(),
        'ahk_id': win.id,
        'x_position': pos.x,
        'y_position': pos.y,
        'height': pos.height,
        'width': pos.width,
    }
    return win_info

@mcp.tool()
async def get_window_text(window_id: str, ctx: Context) -> str:
    """Given a window ID, retrieve the text of the window using the Windows API."""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    return await win.get_text()

# Access type-safe lifespan context in tools
@mcp.tool()
async def get_all_window_info(ctx: Context) -> dict[str, WindowInfo]:
    """Returns information for all (non-hidden) windows, keyed by window ID"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win_list = await ahk.list_windows()
    ret = {}
    for win in win_list:
        win_info = await window_to_info(win)
        ret[win.id] = win_info
    return ret



@mcp.tool()
async def send_keys_to_window(window_id: str, keys: str, ctx: Context) -> None:
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    await win.send(keys)


@mcp.tool()
async def list_window_controls(window_id: str, ctx: Context) -> dict[tuple[str, str], dict[str, Any]]:
    """Given a window id, return information about controls in the window"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    controls = await win.list_controls()
    ret = {}
    for control in controls:
        pos = await control.get_position()
        control_info = {
            'hwnd': control.hwnd,
            'class': control.control_class,
            'x_position': pos.x,
            'y_position': pos.y,
            'height': pos.height,
            'width': pos.width,
            'text': await control.get_text(),
            'window_id': control.window.id,
        }
        ret[(control.hwnd, control.control_class)] = control_info
    return ret

@mcp.tool()
async def send_keys_to_control(window_id: str, control_class: str, keys: str, ctx: Context) -> None:
    """Send keys to a control identified by the window ID and control class"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.control_send(title=f'ahk_id {window_id}', control=control_class, keys=keys)

@mcp.tool()
async def send_keys_to_control_using_hwnd(control_hwnd: str, keys: str, ctx: Context) -> None:
    """Send keys to control using hwnd. May be more reliable when a window has multiple controls of the same class"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.control_send(title=f'ahk_id {control_hwnd}', keys=keys)

@mcp.tool()
async def move_mouse_to_screen_coordinates(x: int, y: int, ctx: Context, speed: int | None = None) -> None:
    """Move the mouse to a position on the screen"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.mouse_move(x, y, speed=speed, coord_mode='Screen')



@mcp.tool()
async def move_mouse_relative(x_offset: int, y_offset: int, ctx: Context, speed: int | None = None) -> None:
    """Move the mouse with offsets relative to its current position"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.mouse_move(x_offset, y_offset, relative=True, speed=speed)


@mcp.tool()
async def get_mouse_position_on_screen(ctx: Context) -> dict[Literal['x', 'y'], int]:
    """Get the mouse position relative to the screen (x,y)"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    x, y = await ahk.get_mouse_position(coord_mode='Screen')
    return {'x': x, 'y': y}

@mcp.tool()
async def get_mouse_position_relative_to_active_window(ctx: Context) -> dict[Literal['x', 'y'], int]:
    """Get the mouse position relative to the active window"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    x, y = await ahk.get_mouse_position(coord_mode='Window')
    return {'x': x, 'y': y}


@mcp.tool()
async def mouse_click(ctx: Context) -> None:
    """Click the left mouse button at its current position"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.click()

@mcp.tool()
async def mouse_click_at_screen_coordinates(x: int, y: int, ctx: Context) -> None:
    """Click the left mouse button at coordinates relative to the screen"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.click(x, y, coord_mode='Screen')

@mcp.tool()
async def mouse_click_at_screen_coordinates(x: int, y: int, ctx: Context) -> None:
    """Click the left mouse button at coordinates relative to the client area"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.click(x, y, coord_mode='Client')

@mcp.tool()
async def right_click(ctx: Context) -> None:
    """Click the right mouse button at its current position"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.click(button=2)

@mcp.tool()
async def right_click_at_screen_coordinates(x: int, y: int, ctx: Context) -> None:
    """Click the right mouse button at coordinates relative to the screen"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.click(x, y, coord_mode='Screen', button=2)

@mcp.tool()
async def activate_window(window_id: str, ctx: Context) -> None:
    """Activates a window, making it the active window"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    await win.activate()

@mcp.tool()
async def set_window_always_on_top(window_id: str, ctx: Context) -> None:
    """Sets window style to make it 'always on top'"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    await win.set_always_on_top('On')


@mcp.tool()
async def disable_window_always_on_top(window_id: str, ctx: Context) -> None:
    """Disables the window 'always on top' style. Has no effect if the style is not set"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    await win.set_always_on_top('Off')

@mcp.tool()
async def send_window_to_top(window_id: str, ctx: Context) -> None:
    """Arranges the window z index to be on top of other windows"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    await win.to_top()


@mcp.tool()
async def send_window_to_bottom(window_id: str, ctx: Context) -> None:
    """Arranges the window z index to be below other windows"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = AsyncWindow(engine=ahk, ahk_id=window_id)
    await win.to_bottom()

@mcp.tool()
async def find_window_by_title(title: str, ctx: Context, exact: bool = False) -> str | None:
    """Attempt to find a window containing the given title. If found, returns its window ID, else None. If `exact` is
    provided, matches exactly on the title"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    win = await ahk.find_window_by_title(title=title, exact=exact)
    if not win:
        return None
    else:
        return win.id

@mcp.tool()
async def get_clipboard_contents(ctx: Context) -> str:
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.get_clipboard()

@mcp.tool()
async def set_clipboard_contents(text_content: str, ctx: Context) -> None:
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    await ahk.set_clipboard(text_content)


@mcp.tool()
async def wait_for_clipboard_contents_to_change(ctx: Context, timeout_seconds: int = 10, any_data: bool = False) -> bool:
    """Wait `timeout_seconds` for the clipboard contents to change with text contents. Returns True if the clipboard change, False if the operation timed out. If `any_data` is True, waits for any content instead of just text content"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    try:
        await ahk.clip_wait(timeout=timeout_seconds, wait_for_any_data=any_data)
        return True
    except TimeoutError:
        return False


class Region(TypedDict):
    left: int
    top: int
    width: int
    height: int

_reader = easyocr.Reader(['en'])

def capture_and_ocr(sct: MSSBase, region: Region) -> str:
    sct_img = sct.grab(region)  # type: ignore[arg-type]
    image_np = np.array(sct_img)
    results = _reader.readtext(image_np, detail=0)
    text = ' '.join(results)
    return text


def detailed_capture_and_ocr(sct: MSSBase, region: Region) -> list[tuple[list[list[int]], str, float]]:
    sct_img = sct.grab(region)  # type: ignore[arg-type]
    image_np = np.array(sct_img)
    return _reader.readtext(image_np)


@mcp.tool()
def ocr_region(region: Region) -> str:
    """Given a region defining a bounding box (relative to the screen), screen captures the region and performs OCR, returning the text found"""
    with mss() as sct:
        return capture_and_ocr(sct, region)

class OcrDetail(TypedDict):
    bounding_box: list[list[int | float]]
    text: str
    confidence: float

@mcp.tool()
def detailed_ocr_region(region: Region) -> list[OcrDetail]:
    """
    Given a region (relative to the screen), screen captures the region and performs OCR returning the
    location of the text found as a bounding box (relative to the capture region), the text identified,
    and a confidence value between 0 and 1. This is useful when you need to know the location of the identified text.
    """
    with mss() as sct:
        result = detailed_capture_and_ocr(sct, region)
    results = [
        {
            'bounding_box': bounding_box,
            'text': text,
            'confidence': confidence
        }
        for bounding_box, text, confidence in result
    ]
    return results

class MonitorInfo(TypedDict):
    name: str
    size: tuple[int, int]
    position: tuple[int, int]
    refresh_rate_millihertz: int | None
    handle: int

def monitor_info_from_monitor(mon: wmutil.Monitor) -> 'MonitorInfo':
    return {
        'name': mon.name,
        'size': mon.size,
        'position': mon.position,
        'refresh_rate_millihertz': mon.refresh_rate_millihertz,
        'handle': mon.handle
    }

@mcp.tool()
def get_monitor_from_point(x: int, y: int) -> MonitorInfo:
    """Given x,y screen coordinates, identifies the monitor at these coordinates and returns its information."""
    mon = wmutil.get_monitor_from_point(x, y)
    return monitor_info_from_monitor(mon)

@mcp.tool()
def get_monitor_of_window(window_id: str) -> MonitorInfo:
    """Given a window ID, identifies the monitor in which the window is located and returns its information."""
    hwnd = int(window_id, 0)
    mon = wmutil.get_window_monitor(hwnd)
    return monitor_info_from_monitor(mon)


@mcp.tool()
def enumerate_monitors() -> list[MonitorInfo]:
    return [monitor_info_from_monitor(mon) for mon in wmutil.enumerate_monitors()]


@mcp.tool()
def get_primary_monitor() -> MonitorInfo:
    return monitor_info_from_monitor(wmutil.get_primary_monitor())

@mcp.tool()
async def save_clipboard_contents(ctx: Context, save_file_path: str) -> None:
    """Save the clipboard contents as a binary blop to a file, specified by `save_file_path`"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    contents = await ahk.get_clipboard_all()
    with open(save_file_path, 'wb') as f:
        f.write(contents)


@mcp.tool()
async def restore_clipboard_contents(ctx: Context, save_file_path: str):
    """Restore previously saved clipboard contents from a file, specified by `save_file_path`"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    with open(save_file_path, 'rb') as f:
        contents = f.read()
    await ahk.set_clipboard_all(contents)


@mcp.tool()
async def wait_for_window(ctx: Context, title: str = '', text: str = '', exclude_title: str = '', exclude_text: str = '', *,
                   title_match_mode: Optional[TitleMatchMode] = None, detect_hidden_windows: Optional[bool] = None,
                   timeout: Optional[int] = 15) -> str | None:
    """Like AutoHotkey's WinWait. Waits for the window and return the window ID if found. If timeout occurs, returns None"""
    ahk: AsyncAHK = ctx.request_context.lifespan_context.ahk
    try:
        win = await ahk.win_wait(title=title, text=text, exclude_title=exclude_title, exclude_text=exclude_text, title_match_mode=title_match_mode, detect_hidden_windows=detect_hidden_windows, timeout=timeout)
        return win.id
    except TimeoutError:
        return None


if __name__ == "__main__":
    # Initialize and run the server
    mcp.run()
