"""
Simple Lightroom keyboard helper for faster working, the toy yet useful app for pywinauto studying

The most obvious usage is to control the exposure/color/tint with the numeric keys.
Have no reasonable explanation why the Lightroom itself lacks the keyboard setup feature out of the box for years

There are the third-party key mappers there, but they are too far from perfection to my taste.
So here is my own *perfect* one

No config files, all the settings are hardcoded

No mature error handling either

Yet it works for me so I see no some tiny reason to improve)

https://habr.com/ru/post/323962/
"""

# class name usage from within class itself
from __future__ import annotations

from pywinauto import Application  # noqa
from pywinauto.win32_hooks import Hook, KeyboardEvent  # noqa
from pywinauto.controls.hwndwrapper import HwndWrapper  # noqa
from pywinauto.application import WindowSpecification  # noqa
from pywinauto.base_wrapper import BaseWrapper  # noqa
from pywinauto.win32structures import RECT  # noqa
from pywinauto.keyboard import send_keys  # noqa
from typing import Optional, Union
import attr
import re
import threading
from queue import Queue
from enum import Enum, unique

# https://stackoverflow.com/questions/25127673/how-to-print-utf-8-to-console-with-python-3-4-windows-8
# UTF8 console output if needed
# sys.stdout.reconfigure(encoding='utf-8')  # noqa

# https://github.com/pywinauto/pywinauto/blob/master/examples/notepad_fast.py
# https://github.com/pywinauto/pywinauto/blob/master/pywinauto/timings.py
# https://pywinauto.readthedocs.io/en/latest/wait_long_operations.html#global-timings-for-all-actions
# Timings.fast()
# Timings.window_find_timeout = 10

# LightRoom main window features
LRMainWindow = Union[BaseWrapper, WindowSpecification]

# emulate click
# see pywinauto.win32_hooks.ID_TO_KEY
CLICKS = {
    'Numpad7': 'BTN_TEMP_MINUS',  # cold
    'Numpad9': 'BTN_TEMP_PLUS',  # warm
    'Numpad4': 'BTN_EXP_MINUS',  # exposure -
    'Numpad6': 'BTN_EXP_PLUS',  # exposure +
    'Numpad1': 'BTN_TINT_MINUS',  # green
    'Numpad3': 'BTN_TINT_PLUS',  # purple
    'Subtract': 'BTN_WHITE_MINUS_BIG',  # white -
    'Add': 'BTN_WHITE_PLUS_BIG',  # white +
    'Divide': 'BTN_CONTRAST_MINUS',  # contrast -
    'Multiply': 'BTN_CONTRAST_PLUS',  # contrast +
    'Numpad0': 'BTN_SHADOW_MINUS_BIG',  # shadow -
    'Decimal': 'BTN_SHADOW_PLUS_BIG',  # shadow +
}

# emulate keypress
KEYS = {
    'Numpad5': "^%v"  # previous settings
}


@unique
class QMsgType(Enum):
    """Queue message type"""
    CLICK = 'click'  # send mouse click
    KEY = 'key'  # send keys
    TERMINATE = 'terminate'  # terminate loop


class HookAlt(Hook):
    """
    Quick and dirty workaround against event altering impossibility

    No error checking or something

    See :py:mod:`pwinauto.win32_hooks`
    """

    def _keyboard_ll_hdl(self, code, event_code, kb_data_ptr):
        if self.handler:
            current_key = self._process_kbd_data(kb_data_ptr)
            event_type = self._process_kbd_msg_type(event_code, current_key)
            event = KeyboardEvent(current_key, event_type, self.pressed_keys)
            if self.handler(event):
                return 1

        old_handler = self.handler
        self.handler = None
        res = super()._keyboard_ll_hdl(code, event_code, kb_data_ptr)
        self.handler = old_handler
        return res


@attr.s(slots=True, auto_attribs=True)
class Lightroom:
    """Adobe Lightroom keyboard helper class"""
    PATH: str = r"C:\Program Files\Adobe\Adobe Lightroom Classic\Lightroom.exe"

    # pwinauto main window object
    WND: Optional[LRMainWindow] = None

    # pwinauto keyboard hook
    KB_HOOK: Optional[HookAlt] = None

    # asyncronous command executor message queue
    QUEUE: Queue = Queue(maxsize=-1)

    # buttons that are recognized
    BTN_TEMP_MINUS_BIG: Optional[HwndWrapper] = None
    BTN_TEMP_MINUS: Optional[HwndWrapper] = None
    BTN_TEMP_PLUS: Optional[HwndWrapper] = None
    BTN_TEMP_PLUS_BIG: Optional[HwndWrapper] = None

    BTN_TINT_MINUS_BIG: Optional[HwndWrapper] = None
    BTN_TINT_MINUS: Optional[HwndWrapper] = None
    BTN_TINT_PLUS: Optional[HwndWrapper] = None
    BTN_TINT_PLUS_BIG: Optional[HwndWrapper] = None

    BTN_EXP_MINUS_BIG: Optional[HwndWrapper] = None
    BTN_EXP_MINUS: Optional[HwndWrapper] = None
    BTN_EXP_PLUS: Optional[HwndWrapper] = None
    BTN_EXP_PLUS_BIG: Optional[HwndWrapper] = None

    BTN_CONTRAST_MINUS_BIG: Optional[HwndWrapper] = None
    BTN_CONTRAST_MINUS: Optional[HwndWrapper] = None
    BTN_CONTRAST_PLUS: Optional[HwndWrapper] = None
    BTN_CONTRAST_PLUS_BIG: Optional[HwndWrapper] = None

    BTN_SHADOW_MINUS_BIG: Optional[HwndWrapper] = None
    BTN_SHADOW_MINUS: Optional[HwndWrapper] = None
    BTN_SHADOW_PLUS: Optional[HwndWrapper] = None
    BTN_SHADOW_PLUS_BIG: Optional[HwndWrapper] = None

    BTN_WHITE_MINUS_BIG: Optional[HwndWrapper] = None
    BTN_WHITE_MINUS: Optional[HwndWrapper] = None
    BTN_WHITE_PLUS: Optional[HwndWrapper] = None
    BTN_WHITE_PLUS_BIG: Optional[HwndWrapper] = None

    @classmethod
    def connect(cls, path: str = None):
        """
        Constructor

        Gets the global keyboard hook and evaluates the control buttons references

        Doesn't return the control to the callee until `disconnect` is called from the keyboard handler

        :param path: path to Lightroom executable
        :return:
        """
        self = cls()

        path = path if path else self.PATH
        # https://pywinauto.readthedocs.io/en/latest/code/pywinauto.application.html#pywinauto.application.Application.connect
        self.WND = Application(backend='win32').connect(path=path).top_window()
        # .connect(title_re="^.* - Adobe Photoshop Lightroom Classic - .*$")
        # self.WND.print_control_identifiers(filename="__123")

        self.bind_buttons()
        self.run_queue_listener()
        self.hook_keyboard()
        return self

    def is_active(self) -> bool:
        """
        Checks if the window active (contains the input focus)

        See also `bring_to_front`

        :return: bool
        """
        wnd = self.WND
        if not wnd:
            return False
        try:
            wnd.wait('active', 0.0001)
            return True
        except:  # noqa
            return False

    def disconnect(self) -> Lightroom:
        """
        Releases the keyboard hook and the button references

        :return:
        """
        self.unhook_keyboard()
        self.QUEUE.put_nowait([QMsgType.TERMINATE])
        for k in self.__slots__:
            if re.match(r'^(?:BTN_|WND)', k):
                setattr(self, k, None)
        return self

    def run_queue_listener(self):
        """
        Asyncronous executor

        Listen the queue and sends the commands received to the Lightroom

        :return: None
        """

        def _listen():
            while True:
                event_type: str
                value: Union[str, HwndWrapper]
                event_type, value = self.QUEUE.get()
                if event_type == QMsgType.CLICK:
                    if not value:
                        continue
                    # Don't ask me why one should use the click+click for I have no explaination,
                    # but only the every 2nd click is working
                    value.click()
                    value.click()
                elif event_type == QMsgType.KEY:
                    send_keys(value)
                elif event_type == QMsgType.TERMINATE:
                    break

        threading.Thread(target=_listen, daemon=True).start()

    def hook_keyboard(self) -> Lightroom:
        def on_event(event) -> bool:
            """
            Keyboard handler body

            Puts the events to the processing queue and returns the control at once,
            otherwise the race condition probability is huge

            See `CLICKS`, `KEYS` and `run_queue_listener`

            :param event:
            :return: True if the event was processed, False otherwise
            """
            if not isinstance(event, KeyboardEvent):
                return False
            if not event.event_type == 'key down':
                return False
            if not self.is_active():
                return False

            # if args.current_key == 'A' and args.event_type == 'key down' and 'Lcontrol' in args.pressed_key:
            #     print("Ctrl + A was pressed")

            key = event.current_key
            if key in CLICKS:
                # put the keypress sequence to the queue and return immediately.
                self.QUEUE.put_nowait([QMsgType.CLICK, getattr(self, CLICKS[key])])
                return True
                # return click(clicks[key])
            elif key in KEYS:
                self.QUEUE.put_nowait([QMsgType.KEY, KEYS[key]])
                return True
                # return send_key(keys[key])

            return False

        # https://github.com/pywinauto/pywinauto/blob/master/examples/hook_and_listen.py
        if self.KB_HOOK is None:
            self.KB_HOOK = HookAlt()
            self.KB_HOOK.handler = on_event
            self.KB_HOOK.hook(keyboard=True, mouse=False)
        return self

    def unhook_keyboard(self) -> Lightroom:
        """
        Release the keyboard global hook

        :return: self
        """
        if self.KB_HOOK is not None:
            self.KB_HOOK.unhook_keyboard()
            self.KB_HOOK = None
        return self

    def bring_to_front(self) -> LRMainWindow:
        """
        Bring window topmost and get the input focus

        :return: main window reference
        """
        wnd = self.ensure_hwnd()
        wnd.set_focus()
        # https://stackoverflow.com/questions/47494592/bring-an-application-window-gui-in-front-on-disktop-python3-tkinter
        # from pywinauto.findwindows import find_window
        # from pywinauto.win32functions import SetForegroundWindow
        # SetForegroundWindow(find_window(title='gui title'))
        return wnd

    def ensure_hwnd(self) -> LRMainWindow:
        """
        Checks if the control was binded already to the property

        :return: control reference or throw exception
        """
        wnd = self.WND
        if wnd is None:
            raise Exception('HWND was not assigned yet')
        return wnd

    def bind_buttons(self) -> Lightroom:
        """
        Bind the Lightroom buttons to the instance properties

        The coordinate numbers were gotten with the AutoIt's `Au3Info` on the 1920x1080 screen

        The `tone_control` button is the "base" because it's name is unique

        It's top left corner was 1656, 404, so we subtract this from the
        other button's coordinates and add 3 just not to click to the very angle

        :return: self
        """

        def set_btn(name: str, spy_x: int, spy_y: int) -> HwndWrapper:
            """
            Find the actual button in the coordinates and assigns it's reference
            to the related property or raises if something is wrong

            The textual property name is choosen instead of the property itself
            for the error messaging purpose as there's no stright way to get the property name from it's reference
            while getting the property value by name is trivial

            It's not all good as we must ensure strings are matched to the properties and vice versa

            :param name: property name
            :param spy_x: X coordinate (left)
            :param spy_y: Y coordinate (top)
            :return: the button reference
            """

            # noinspection PyTypeChecker
            wrapper: HwndWrapper = wnd.from_point(r.left + spy_x - 1656 + 3, r.top + spy_y - 404 + 3)

            # for example, the buttons can not be read if no photos are loaded
            # the clickable "Button"s turn into the disabled "View"s then
            if wrapper.element_info.class_name != 'Button' or wrapper.element_info.rich_text != '':
                raise Exception(f"Can't find button {name} at pos {spy_x}:{spy_y}")
            setattr(self, name, wrapper)
            return wrapper

        wnd = self.bring_to_front()
        # https://pywinauto.readthedocs.io/en/latest/code/pywinauto.win32_element_info.html?highlight=control_type#pywinauto-win32-element-info
        # https://github.com/pywinauto/pywinauto/blob/master/pywinauto/win32_element_info.py -> children
        # https://stackoverflow.com/questions/45140434/find-new-window-dialogs-automatically
        tone_control = wnd.descendants(class_name='Static', title='Tone Control')[0]
        r: RECT = tone_control.element_info.rectangle

        # point = tone_control.button(0).rectangle().mid_point()
        # screen_r = tone_control.client_to_screen((r.left, r.top))
        # print(f"{info.control_id=}, {info.handle=:x}, {info.rectangle=}")
        # btns = get_control_from_coord(all_btns, r.left + 1748 - 1656, r.top + 433 - 404)
        # btn = dlg.from_point(screen_r[0] + 1748 - 1656, screen_r[1] + 433 - 404)

        set_btn('BTN_TEMP_MINUS_BIG', 1748, 344)
        set_btn('BTN_TEMP_MINUS', 1779, 344)
        set_btn('BTN_TEMP_PLUS', 1807, 344)
        set_btn('BTN_TEMP_PLUS_BIG', 1836, 344)

        set_btn('BTN_TINT_MINUS_BIG', 1748, 369)
        set_btn('BTN_TINT_MINUS', 1779, 369)
        set_btn('BTN_TINT_PLUS', 1807, 369)
        set_btn('BTN_TINT_PLUS_BIG', 1836, 369)

        set_btn('BTN_EXP_MINUS_BIG', 1748, 433)
        set_btn('BTN_EXP_MINUS', 1779, 433)
        set_btn('BTN_EXP_PLUS', 1807, 433)
        set_btn('BTN_EXP_PLUS_BIG', 1836, 433)

        set_btn('BTN_CONTRAST_MINUS_BIG', 1748, 458)
        set_btn('BTN_CONTRAST_MINUS', 1779, 458)
        set_btn('BTN_CONTRAST_PLUS', 1807, 458)
        set_btn('BTN_CONTRAST_PLUS_BIG', 1836, 458)

        set_btn('BTN_SHADOW_MINUS_BIG', 1748, 516)
        set_btn('BTN_SHADOW_MINUS', 1779, 516)
        set_btn('BTN_SHADOW_PLUS', 1807, 516)
        set_btn('BTN_SHADOW_PLUS_BIG', 1836, 516)

        set_btn('BTN_WHITE_MINUS_BIG', 1748, 541)
        set_btn('BTN_WHITE_MINUS', 1779, 541)
        set_btn('BTN_WHITE_PLUS', 1807, 541)
        set_btn('BTN_WHITE_PLUS_BIG', 1836, 541)

        return self

    def __del__(self):
        self.disconnect()


LR = Lightroom.connect()
