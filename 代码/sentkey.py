import serial
import threading
import time
import ctypes
from ctypes import wintypes


# 允许的字符集合
ALLOWED_CHARS = set("abcdefghijklmnopqrstuvwxyz0123456789-=[]\\;',./")


class ArduinoOutput:
    """Arduino输出，支持字母、数字和额外符号"""
    
    def __init__(self, serial_port: serial.Serial):
        self.ser = serial_port
        self._lock = threading.Lock()
        self._pressed = set()
        self._alive = True
    
    def _send(self, data: bytes):
        with self._lock:
            if not self._alive or not self.ser or not self.ser.is_open:
                return False
            try:
                self.ser.write(data)
                time.sleep(0.0001)
                return True
            except Exception as e:
                print(f"[ArduinoOutput] 写入失败: {e}")
                self._alive = False
                return False
    
    def press(self, key: str):
        if not self._alive:
            return
        c = key[0] if key else ''
        if c.isalpha():
            c = c.lower()
        if c not in ALLOWED_CHARS:
            return
        
        if c not in self._pressed:
            self._pressed.add(c)
            self._send(f"K:{c}\n".encode())
    
    def release(self, key: str):
        if not self._alive:
            return
        c = key[0] if key else ''
        if c.isalpha():
            c = c.lower()
        if c not in ALLOWED_CHARS:
            return
            
        if c in self._pressed:
            self._pressed.discard(c)
            self._send(f"R:{c}\n".encode())
    
    def release_all(self):
        """安全释放所有按键"""
        if not self._alive:
            self._pressed.clear()
            return
        
        with self._lock:
            keys = list(self._pressed)
            self._pressed.clear()
            
            for c in keys:
                try:
                    if self.ser and self.ser.is_open:
                        self.ser.write(f"R:{c}\n".encode())
                except Exception as e:
                    print(f"[ArduinoOutput] release_all 写入失败: {e}")
                    self._alive = False
                    break
            
            try:
                if self.ser and self.ser.is_open:
                    self.ser.write(b"X\n")
            except Exception as e:
                print(f"[ArduinoOutput] 发送X命令失败: {e}")
                self._alive = False


KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_SCANCODE = 0x0008


class KeyBdInput(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]


class HardwareInput(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD)
    ]


class MouseInput(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]


class InputUnion(ctypes.Union):
    _fields_ = [
        ("mi", MouseInput),
        ("ki", KeyBdInput),
        ("hi", HardwareInput)
    ]


class Input(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("union", InputUnion)
    ]


INPUT_KEYBOARD = 1


class VirtualOutput:
    """Windows API 虚拟按键输出 - 支持字母、数字和额外符号"""
    
    def __init__(self, debug=False):
        self._lock = threading.Lock()
        self._pressed = set()
        self._alive = True
        self._debug = debug
        
        self._vk_map = self._build_vk_map()
    
    def _build_vk_map(self):
        """构建字符到虚拟键码的映射"""
        vk_map = {}
        
        for i in range(26):
            char = chr(ord('a') + i)
            vk = 0x41 + i
            vk_map[char] = vk
        
        for i in range(10):
            char = str(i)
            vk = 0x30 + i
            vk_map[char] = vk
        
        vk_map['-'] = 0xBD
        vk_map['='] = 0xBB
        vk_map['['] = 0xDB
        vk_map[']'] = 0xDD
        vk_map['\\'] = 0xDC
        vk_map[';'] = 0xBA
        vk_map["'"] = 0xDE
        vk_map[','] = 0xBC
        vk_map['.'] = 0xBE
        vk_map['/'] = 0xBF
        
        return vk_map
    
    def _send_key(self, vk, is_press):
        with self._lock:
            if not self._alive:
                return False
            try:
                extra = ctypes.c_ulong(0)
                union = InputUnion()
                
                union.ki = KeyBdInput(
                    wVk=vk,
                    wScan=0,
                    dwFlags=0 if is_press else KEYEVENTF_KEYUP,
                    time=0,
                    dwExtraInfo=ctypes.pointer(extra)
                )
                
                x = Input(type=INPUT_KEYBOARD, union=union)
                ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))
                return True
            except Exception as e:
                print(f"[VirtualOutput] 发送按键失败: {e}")
                return False
    
    def press(self, key: str):
        if not self._alive:
            return
        c = key[0] if key else ''
        if c.isalpha():
            c = c.lower()
        if c not in ALLOWED_CHARS:
            return
        
        vk = self._vk_map.get(c)
        if vk and c not in self._pressed:
            self._pressed.add(c)
            if self._debug:
                print(f"[VirtualOutput] PRESS: key={c}, vk=0x{vk:02X}")
            self._send_key(vk, True)
    
    def release(self, key: str):
        if not self._alive:
            return
        c = key[0] if key else ''
        if c.isalpha():
            c = c.lower()
        if c not in ALLOWED_CHARS:
            return
            
        vk = self._vk_map.get(c)
        if vk and c in self._pressed:
            self._pressed.discard(c)
            if self._debug:
                print(f"[VirtualOutput] RELEASE: key={c}, vk=0x{vk:02X}")
            self._send_key(vk, False)
    
    def release_all(self):
        """安全释放所有按键"""
        if not self._alive:
            self._pressed.clear()
            return
        
        with self._lock:
            keys = list(self._pressed)
            self._pressed.clear()
            
            for c in keys:
                try:
                    vk = self._vk_map.get(c)
                    if vk:
                        extra = ctypes.c_ulong(0)
                        union = InputUnion()
                        union.ki = KeyBdInput(
                            wVk=vk,
                            wScan=0,
                            dwFlags=KEYEVENTF_KEYUP,
                            time=0,
                            dwExtraInfo=ctypes.pointer(extra)
                        )
                        x = Input(type=INPUT_KEYBOARD, union=union)
                        ctypes.windll.user32.SendInput(1, ctypes.pointer(x), ctypes.sizeof(x))
                except Exception as e:
                    print(f"[VirtualOutput] release_all 释放失败: {e}")
                    break
