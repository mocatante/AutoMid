import sys
import os
import shutil
import configparser
import struct
import time
import threading
import ctypes
import glob
import serial
import serial.tools.list_ports
from sentkey import ArduinoOutput, VirtualOutput, ALLOWED_CHARS
from ctypes import wintypes
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QLineEdit,
    QPushButton, QComboBox, QVBoxLayout, QHBoxLayout,
    QFileDialog, QMessageBox, QCheckBox, QTabWidget, QSizePolicy
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QThread
from PyQt5.QtGui import QFont, QDoubleValidator, QIntValidator

# Windows MIDI API 常量
MIDI_MAPPER = -1
CALLBACK_NULL = 0


def get_base_dir():
    """获取程序运行根目录（打包后返回exe所在目录）"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.abspath(__file__))


def get_pitch_map(ini_filename="map.ini", section="PitchMap"):
    """
    读取map.ini，构建数字音高→字符的映射字典
    优先从exe同级目录读取，失败则尝试多个备选路径
    返回 (pitch_map, status_message)
    """
    base_dir = get_base_dir()
    ini_path = os.path.join(base_dir, ini_filename)
    
    # 备选路径列表（按优先级）
    possible_paths = [
        ini_path,  # 首选：exe/脚本所在目录
        os.path.join(os.getcwd(), ini_filename),  # 当前工作目录
        os.path.join(os.path.expanduser("~"), ini_filename),  # 用户主目录
    ]
    
    # 开发环境额外路径
    if not getattr(sys, 'frozen', False):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        possible_paths.extend([
            os.path.join(script_dir, ini_filename),
            os.path.join(script_dir, "..", ini_filename),
        ])
    
    cp = configparser.ConfigParser()
    found_path = None
    
    for path in possible_paths:
        if os.path.exists(path):
            try:
                cp.read(path, encoding="utf-8")
                found_path = path
                status_message = f"✅ 已加载映射文件: {os.path.basename(path)}"
                print(status_message)
                break
            except Exception as e:
                print(f"⚠️ 读取失败 {path}: {e}")
                continue
    
    if not found_path:
        searched = "\n  ".join(possible_paths)
        raise FileNotFoundError(f"找不到 {ini_filename}，已尝试:\n  {searched}")

    if section not in cp.sections():
        raise KeyError(f"ini文件中未找到 [{section}] 节，请检查格式")

    pitch_map = {}
    for num_str, char in cp[section].items():
        num_str = num_str.replace(":", "")
        try:
            pitch_num = int(num_str)
        except ValueError:
            raise ValueError(f"[{section}] 中的键 '{num_str}' 不是有效数字")
        pitch_char = char.strip()
        if pitch_char:
            pitch_map[pitch_num] = pitch_char

    if not pitch_map:
        raise ValueError(f"[{section}] 中无有效映射")

    status_message += f"\n✅ 成功加载 {len(pitch_map)} 个按键映射"
    print(f"✅ 成功加载 {len(pitch_map)} 个按键映射")
    return pitch_map, status_message


def read_midi(mid_file_path):
    with open(mid_file_path, 'rb') as f:
        mid_buf = f.read()
        print(f"✅ MID文件读取成功, 共 {len(mid_buf)} 字节")
    return mid_buf


def read_vlq(buf, offset):
    """读取MIDI变长值(Variable Length Quantity)"""
    value = 0
    while True:
        b = buf[offset]
        value = (value << 7) | (b & 0x7F)
        offset += 1
        if not (b & 0x80):
            break
    return value, offset


def read_mid_note(buf, point, running_status, track_end):
    """读取单个MIDI事件"""
    if point >= track_end:
        raise ValueError("读取位置已超出轨道范围")
    
    delta, point = read_vlq(buf, point)
    
    if point >= track_end:
        raise ValueError("读取状态字节时超出轨道范围")
    
    status = buf[point]
    if status & 0x80:
        running_status = status
        point += 1
    
    events = []
    is_end = False
    high_nibble = running_status & 0xF0
    
    if high_nibble in (0x80, 0x90, 0xA0, 0xB0, 0xE0):
        if point + 2 > track_end:
            raise ValueError(f"通道消息数据不足")
        data1 = buf[point]
        data2 = buf[point + 1]
        point += 2
        
        if high_nibble == 0x80:
            events.append({'type': 'note_off', 'pitch': data1, 'vel': data2})
        elif high_nibble == 0x90:
            if data2 == 0:
                events.append({'type': 'note_off', 'pitch': data1, 'vel': data2})
            else:
                events.append({'type': 'note_on', 'pitch': data1, 'vel': data2})
                
    elif high_nibble in (0xC0, 0xD0):
        if point + 1 > track_end:
            raise ValueError(f"程序改变消息数据不足")
        point += 1
        
    elif running_status == 0xFF:
        if point >= track_end:
            raise ValueError("读取Meta事件类型时超出范围")
        meta_type = buf[point]
        point += 1
        
        length, point = read_vlq(buf, point)
        
        if point + length > track_end:
            raise ValueError(f"Meta事件数据长度不足")
        
        data = buf[point:point+length]
        point += length
        
        if meta_type == 0x2F:
            is_end = True
        elif meta_type == 0x51 and length >= 3:
            tempo = (data[0] << 16) | (data[1] << 8) | data[2]
            events.append({'type': 'tempo', 'value': tempo})
            
    elif running_status in (0xF0, 0xF7):
        length, point = read_vlq(buf, point)
        point += length
        
    return delta, events, point, running_status, is_end


def format_time(t):
    """将时间转换为字符串（去除6位限制）"""
    return str(int(t))


def parse_midi(name, mid_buf, pitch_map, speed=1.0):
    """解析MIDI并生成移调后的txt文件"""
    if len(mid_buf) < 14:
        raise ValueError("MIDI文件过短，无法解析")
    
    if mid_buf[0:4] != b'MThd':
        raise ValueError("无效的MIDI文件：缺少MThd头标记")
    
    fmt, tracks_count, division = struct.unpack('>HHH', mid_buf[8:14])
    
    if division & 0x8000:
        mode = 'SMPTE'
        fps_byte = (division >> 8) & 0xFF
        fps = -struct.unpack('b', bytes([fps_byte]))[0]
        ticks_per_frame = division & 0xFF
        if fps == 0 or ticks_per_frame == 0:
            raise ValueError("SMPTE模式下FPS或TicksPerFrame不能为0")
        tpq = None
        print(f"🎬 SMPTE模式: {fps}fps, {ticks_per_frame}ticks/frame")
    else:
        mode = 'TPQ'
        tpq = division & 0x7FFF
        if tpq == 0:
            raise ValueError("TPQ模式下TicksPerQuarter不能为0")
        fps = None
        ticks_per_frame = None
        print(f"🎵 TPQ模式: {tpq} ticks/quarter")
    
    file_out = []
    base_dir = get_base_dir()
    temp_dir = os.path.join(base_dir, "temp")
    os.makedirs(temp_dir, exist_ok=True)
    
    for transpose in range(-24, +25):
        all_events = []
        all_notes_count = 0
        invalid_notes_count = 0
        
        offset = 14
        
        for t in range(tracks_count):
            while offset < len(mid_buf) - 4:
                if mid_buf[offset:offset+4] == b'MTrk':
                    break
                offset += 1
            else:
                raise ValueError(f"轨道{t+1}未找到MTrk标记")
            
            offset += 4
            track_len = struct.unpack('>I', mid_buf[offset:offset+4])[0]
            offset += 4
            
            track_end = offset + track_len
            if track_end > len(mid_buf):
                raise ValueError(f"轨道{t+1}声明长度超出文件范围")
            
            track_events = []
            cur_time = 0
            running_status = 0
            tempo_us = 500000
            
            while offset < track_end:
                delta, events, offset, running_status, is_end = read_mid_note(
                    mid_buf, offset, running_status, track_end
                )
                
                if mode == 'TPQ':
                    time_inc = (delta * tempo_us) / (tpq * 1000.0)
                else:
                    time_inc = (delta * 1000.0) / (ticks_per_frame * fps)
                
                if speed != 1.0:
                    time_inc = time_inc / speed
                
                cur_time += time_inc
                
                for evt in events:
                    if evt['type'] == 'tempo':
                        tempo_us = evt['value']
                    elif evt['type'] in ('note_on', 'note_off'):
                        all_notes_count += 1
                        
                        transposed_pitch = evt['pitch'] + transpose
                        if transposed_pitch in pitch_map:
                            char = pitch_map[transposed_pitch]
                            event_type = 'P' if evt['type'] == 'note_on' else 'R'
                            track_events.append({
                                'type': event_type,
                                'key': char,
                                'time': int(cur_time)
                            })
                        else:
                            invalid_notes_count += 1
                
                if is_end:
                    break
            
            all_events.extend(track_events)
        
        all_events.sort(key=lambda x: x['time'])
        
        out_buf = []
        for evt in all_events:
            time_str = format_time(evt['time'])
            line = f"{evt['type']}\t{evt['key']}\t{time_str}"
            out_buf.append(line)
        
        if all_notes_count == 0:
            valid_percent = 0.0
        else:
            valid_percent = round((all_notes_count - invalid_notes_count) / all_notes_count * 100, 2)
        
        percent_str = f"{valid_percent:.2f}"
        file_name = f"[{transpose}]{name}[{percent_str}].txt"
        file_path = os.path.join(temp_dir, file_name)
        
        with open(file_path, "w", encoding="utf-8") as f:
            f.write("\n".join(out_buf))
        
        file_out.append({
            "valid_percent": valid_percent,
            "file_name": file_name,
            "file_path": file_path,
        })
    
    file_out.sort(key=lambda x: x["valid_percent"], reverse=True)
    return file_out


class MidiPlayer:
    """Windows MIDI API 播放器"""
    
    def __init__(self):
        self.hMidiOut = wintypes.HANDLE()
        self.is_playing = False
        self.is_paused = False
        self.current_thread = None
        self.events = []
        self.start_time = 0
        self.pause_time = 0
        self.mapping = {}
        self.speed = 1.0
        
    def load_mapping(self, ini_path="map.ini"):
        """加载map.ini并创建反向映射（字符->音高）"""
        base_dir = get_base_dir()
        
        # 备选路径（按优先级）
        possible_paths = [
            os.path.join(base_dir, ini_path),  # exe/脚本同级目录（首选）
            os.path.join(os.getcwd(), ini_path),  # 当前工作目录
        ]
        
        # 开发环境额外路径
        if not getattr(sys, 'frozen', False):
            script_dir = os.path.dirname(os.path.abspath(__file__))
            possible_paths.extend([
                os.path.join(script_dir, ini_path),
                os.path.join(script_dir, "..", ini_path),
                ini_path,
            ])
        
        cp = configparser.ConfigParser()
        found_path = None
        
        for path in possible_paths:
            if os.path.exists(path):
                try:
                    cp.read(path, encoding="utf-8")
                    found_path = path
                    print(f"✅ 已加载映射文件: {path}")
                    break
                except Exception as e:
                    print(f"⚠️ 读取失败 {path}: {e}")
                    continue
        
        if not found_path:
            searched = "\n  ".join(possible_paths)
            raise FileNotFoundError(f"找不到 {ini_path}，已尝试:\n  {searched}")

        mapping = {}
        if "PitchMap" in cp.sections():
            for num_str, char in cp["PitchMap"].items():
                num_str = num_str.replace(":", "")
                try:
                    pitch = int(num_str)
                    char = char.strip()
                    if char:
                        mapping[char] = pitch
                except ValueError:
                    continue
        
        if not mapping:
            raise ValueError("map.ini 中没有找到有效的 PitchMap 映射")
            
        self.mapping = mapping
        print(f"✅ 成功加载 {len(mapping)} 个按键映射")
        return mapping
    
    def parse_txt(self, txt_path):
        """解析生成的txt文件"""
        if not os.path.exists(txt_path):
            raise FileNotFoundError(f"文件不存在: {txt_path}")
            
        events = []
        with open(txt_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or '\t' not in line:
                    continue
                parts = line.split('\t')
                if len(parts) >= 3:
                    evt_type, key, time_str = parts[0], parts[1], parts[2]
                    if key not in self.mapping:
                        continue
                    
                    pitch = self.mapping[key]
                    try:
                        time_val = int(time_str)
                    except ValueError:
                        continue
                    
                    event = {
                        'type': 'on' if evt_type == 'P' else 'off',
                        'pitch': pitch,
                        'time': time_val,
                        'raw_type': evt_type
                    }
                    events.append(event)
        
        events.sort(key=lambda x: x['time'])
        print(f"✅ 解析到 {len(events)} 个事件")
        return events
    
    def init_midi(self):
        """初始化MIDI输出"""
        if self.hMidiOut.value is None or self.hMidiOut.value == 0:
            result = ctypes.windll.winmm.midiOutOpen(
                ctypes.byref(self.hMidiOut),
                MIDI_MAPPER,
                0,
                0,
                CALLBACK_NULL
            )
            if result != 0:
                raise Exception(f"MIDI设备初始化失败，错误码: {result}")
            print("✅ MIDI设备已初始化")
    
    def close_midi(self):
        """关闭MIDI设备"""
        if self.hMidiOut.value and self.hMidiOut.value != 0:
            ctypes.windll.winmm.midiOutClose(self.hMidiOut)
            self.hMidiOut = wintypes.HANDLE()
            print("✅ MIDI设备已关闭")
    
    def send_note(self, pitch, velocity, on=True):
        """发送MIDI音符消息"""
        if not self.hMidiOut.value or self.hMidiOut.value == 0:
            return
        
        status = 0x90 if on else 0x80
        channel = 0
        msg = (status | channel) | (pitch << 8) | (velocity << 16)
        ctypes.windll.winmm.midiOutShortMsg(self.hMidiOut, msg)
    
    def _play_thread(self):
        """播放线程"""
        try:
            self.init_midi()
            self.is_playing = True
            self.start_time = time.time() * 1000
            
            for pitch in range(128):
                self.send_note(pitch, 0, False)
            
            event_index = 0
            total_events = len(self.events)
            
            print(f"▶️ 开始播放，共 {total_events} 个事件")
            
            while self.is_playing and event_index < total_events:
                if self.is_paused:
                    time.sleep(0.01)
                    continue
                
                current_time = (time.time() * 1000 - self.start_time) * self.speed
                event = self.events[event_index]
                
                if current_time >= event['time']:
                    if event['type'] == 'on':
                        self.send_note(event['pitch'], 100, True)
                    else:
                        self.send_note(event['pitch'], 0, False)
                    event_index += 1
                else:
                    time.sleep(0.001)
            
            print("⏹️ 播放结束")
            
        except Exception as e:
            print(f"❌ 播放线程错误: {e}")
        finally:
            for pitch in range(128):
                self.send_note(pitch, 0, False)
            
            self.is_playing = False
            self.is_paused = False
            self.close_midi()
    
    def play(self, txt_path, speed=1.0):
        """开始播放"""
        if self.is_playing:
            self.stop()
        
        if not self.mapping:
            self.load_mapping()
        
        self.events = self.parse_txt(txt_path)
        if not self.events:
            print("❌ 没有找到可播放的音符事件")
            return False
        
        self.speed = speed
        self.current_thread = threading.Thread(target=self._play_thread)
        self.current_thread.daemon = True
        self.current_thread.start()
        return True
    
    def pause(self):
        """暂停/继续切换"""
        if self.is_playing:
            if self.is_paused:
                pause_duration = time.time() * 1000 - self.pause_time
                self.start_time += pause_duration
                self.is_paused = False
                print("▶️ 继续播放")
            else:
                self.is_paused = True
                self.pause_time = time.time() * 1000
                print("⏸️ 暂停播放")
            return True
        return False
    
    def stop(self):
        """停止播放"""
        self.is_playing = False
        self.is_paused = False
        if self.current_thread and self.current_thread.is_alive():
            self.current_thread.join(timeout=1.0)
        self.close_midi()
    
    def is_finished(self):
        """检查是否播放完成"""
        return not self.is_playing and not (self.current_thread and self.current_thread.is_alive())


class AutoMidWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.player = MidiPlayer()
        self.current_mid_path = None
        self.parsed_files = []
        self.base_dir = get_base_dir()
        self.temp_dir = os.path.join(self.base_dir, "temp")
        self.out_dir = os.path.join(self.base_dir, "out")
        self.current_speed = 1.0
        
        os.makedirs(self.out_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)
        
        self.speed_timer = QTimer()
        self.speed_timer.timeout.connect(self.on_speed_apply)
        self.speed_timer.setSingleShot(True)
        
        self.init_ui()
        
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.check_playback_status)
        self.status_timer.start(500)

    def init_ui(self):
        self.setWindowTitle("AutoMid")
        self.setFixedSize(660, 240)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)

        self.drag_label = QLabel("拖拽MID文件到此处，或点击打开文件")
        self.drag_label.setStyleSheet("""
            QLabel {
                border: 2px dashed #666;
                border-radius: 4px;
                color: #333;
                background-color: #f9f9f9;
            }
            QLabel:hover { border-color: #2980b9; background-color: #f0f8ff; }
        """)
        self.drag_label.setAlignment(Qt.AlignCenter)
        self.drag_label.setFont(QFont("微软雅黑", 10))
        self.drag_label.setMinimumHeight(120)
        self.drag_label.setAcceptDrops(True)
        self.drag_label.mousePressEvent = self.on_click_open_file
        self.drag_label.dragEnterEvent = self.label_dragEnterEvent
        self.drag_label.dropEvent = self.label_dropEvent
        main_layout.addWidget(self.drag_label, stretch=3)

        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(8)

        left_layout = QVBoxLayout()
        left_layout.setSpacing(2)
        
        combo_label = QLabel("选择移调方案:")
        combo_label.setFont(QFont("微软雅黑", 9))
        left_layout.addWidget(combo_label)
        
        self.num_combobox = QComboBox()
        self.num_combobox.setFont(QFont("微软雅黑", 9))
        self.num_combobox.setMinimumWidth(200)
        self.num_combobox.setToolTip("按有效率排序，选择最适合键盘演奏的移调方案")
        left_layout.addWidget(self.num_combobox)
        
        bottom_layout.addLayout(left_layout, stretch=3)

        right_layout = QHBoxLayout()
        right_layout.setSpacing(6)
        
        speed_label = QLabel("倍速：")
        speed_label.setFont(QFont("微软雅黑", 9))
        
        self.speed_edit = QLineEdit()
        self.speed_edit.setFixedWidth(60)
        self.speed_edit.setFont(QFont("微软雅黑", 9))
        self.speed_edit.setPlaceholderText("1.0")
        self.speed_edit.setText("1.0")
        
        from PyQt5.QtGui import QDoubleValidator
        self.validator = QDoubleValidator(0.1, 10.0, 2)
        self.validator.setNotation(QDoubleValidator.StandardNotation)
        self.speed_edit.setValidator(self.validator)

        self.listen_btn = QPushButton("试听")
        self.export_btn = QPushButton("导出")
        self.listen_btn.setFixedSize(60, 28)
        self.export_btn.setFixedSize(60, 28)
        self.listen_btn.setFont(QFont("微软雅黑", 9))
        self.export_btn.setFont(QFont("微软雅黑", 9))

        self.listen_btn.clicked.connect(self.on_listen)
        self.export_btn.clicked.connect(self.on_export)
        self.speed_edit.textChanged.connect(self.on_speed_changed)
        
        self.listen_btn.setEnabled(False)
        self.export_btn.setEnabled(False)

        right_layout.addWidget(speed_label)
        right_layout.addWidget(self.speed_edit)
        right_layout.addStretch(1)
        right_layout.addWidget(self.listen_btn)
        right_layout.addWidget(self.export_btn)
        
        bottom_layout.addLayout(right_layout, stretch=2)
        main_layout.addLayout(bottom_layout, stretch=1)

    def toggle_always_on_top(self, state):
        """切换窗口顶置状态"""
        if state == Qt.Checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        self.show()

    def on_speed_changed(self, text):
        if not self.current_mid_path:
            return
            
        try:
            new_speed = float(text) if text else 1.0
            if 0.1 <= new_speed <= 10.0:
                self.speed_timer.start(500)
        except ValueError:
            pass

    def on_speed_apply(self):
        if not self.current_mid_path:
            return
            
        try:
            new_speed = float(self.speed_edit.text() or "1.0")
            if new_speed < 0.1 or new_speed > 10.0:
                return
        except ValueError:
            return
        
        if abs(new_speed - self.current_speed) > 0.001:
            self.regenerate_temp_files(new_speed)

    def regenerate_temp_files(self, speed):
        if self.player.is_playing:
            self.player.stop()
            self.listen_btn.setText("试听")
        
        self.current_speed = speed
        
        current_transpose = None
        if self.num_combobox.currentIndex() >= 0:
            current_text = self.num_combobox.currentText()
            try:
                transpose_str = current_text.split(']')[0][1:]
                current_transpose = transpose_str
            except:
                current_transpose = None
        
        file_name = os.path.basename(self.current_mid_path)
        self.drag_label.setText(f"已选择：{file_name}\n正在重新生成（倍速: {speed}x）...")
        QApplication.processEvents()
        
        try:
            pitch_mapping, status_message = get_pitch_map()
            mid_buf = read_midi(self.current_mid_path)
            base_name = os.path.splitext(file_name)[0]
            
            self.parsed_files = parse_midi(base_name, mid_buf, pitch_mapping, speed)
            
            if self.parsed_files:
                self.fill_combobox_with_results()
                
                if current_transpose:
                    target_str = f"[{current_transpose}]"
                    for i in range(self.num_combobox.count()):
                        if target_str in self.num_combobox.itemText(i):
                            self.num_combobox.setCurrentIndex(i)
                            break
                
                self.drag_label.setText(f"已选择：{file_name}\n{status_message}\n已生成 {len(self.parsed_files)} 个移调方案（倍速: {speed}x）")
                
        except Exception as e:
            QMessageBox.critical(self, "重新生成失败", f"倍速 {speed}x 生成失败：{str(e)}")

    def label_dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and len(event.mimeData().urls()) == 1:
            event.acceptProposedAction()
        else:
            event.ignore()

    def label_dropEvent(self, event):
        file_url = event.mimeData().urls()[0]
        if file_url.isLocalFile():
            self.process_mid_file(file_url.toLocalFile())

    def on_click_open_file(self, event):
        if event.button() == Qt.LeftButton:
            file_path, _ = QFileDialog.getOpenFileName(
                self, "选择MIDI文件", "", 
                "MIDI文件 (*.mid *.midi);;所有文件 (*)"
            )
            if file_path:
                self.process_mid_file(file_path)

    def process_mid_file(self, file_path):
        if not os.path.exists(file_path):
            QMessageBox.critical(self, "错误", "文件不存在")
            return
        
        if self.player.is_playing:
            self.player.stop()
            self.listen_btn.setText("试听")
        
        self.current_mid_path = file_path
        file_name = os.path.basename(file_path)
        self.drag_label.setText(f"已选择：{file_name}\n正在解析...")
        QApplication.processEvents()
        
        try:
            pitch_mapping, status_message = get_pitch_map()
            mid_buf = read_midi(file_path)
            
            speed_input = self.speed_edit.text().strip()
            self.current_speed = float(speed_input) if speed_input else 1.0
            
            base_name = os.path.splitext(file_name)[0]
            self.parsed_files = parse_midi(base_name, mid_buf, pitch_mapping, self.current_speed)
            
            if not self.parsed_files:
                QMessageBox.warning(self, "警告", "没有生成有效的移调文件")
                return
            
            self.fill_combobox_with_results()
            
            self.drag_label.setText(f"已选择：{file_name}\n{status_message}\n已生成 {len(self.parsed_files)} 个移调方案（倍速: {self.current_speed}x）")
            
            self.listen_btn.setEnabled(True)
            self.export_btn.setEnabled(True)
            
        except Exception as e:
            self.drag_label.setText(f"已选择：{file_name}\n解析失败")
            QMessageBox.critical(self, "解析错误", str(e))

    def fill_combobox_with_results(self):
        self.num_combobox.clear()
        
        sorted_files = sorted(self.parsed_files, key=lambda x: x['valid_percent'], reverse=True)
        
        for item in sorted_files:
            try:
                transpose = item['file_name'].split(']')[0][1:]
            except:
                transpose = "?"
            
            percent = item['valid_percent']
            sign = "+" if int(transpose) >= 0 else ""
            display_text = f"[{sign}{transpose}] {percent:.2f}%"
            
            if percent >= 90:
                display_text += " ★推荐"
            elif percent < 50:
                display_text += " (效果较差)"
            
            self.num_combobox.addItem(display_text, item['file_name'])

    def on_listen(self):
        if not self.parsed_files:
            QMessageBox.warning(self, "提示", "请先选择MIDI文件")
            return
        
        if self.listen_btn.text() == "结束":
            self.player.stop()
            self.listen_btn.setText("试听")
            return
        
        index = self.num_combobox.currentIndex()
        if index < 0:
            QMessageBox.warning(self, "提示", "请选择一个移调方案")
            return
        
        file_name = self.num_combobox.itemData(index)
        file_path = os.path.join(self.temp_dir, file_name)
        
        if not os.path.exists(file_path):
            QMessageBox.critical(self, "错误", f"文件不存在: {file_path}\n请检查倍速修改后是否已重新生成")
            return
        
        try:
            speed = float(self.speed_edit.text() or "1.0")
        except ValueError:
            speed = 1.0
        
        if self.player.is_playing:
            self.player.stop()
        
        if self.player.play(file_path, speed):
            self.listen_btn.setText("结束")
        else:
            QMessageBox.warning(self, "错误", "播放失败")

    def check_playback_status(self):
        if self.player.is_finished() and self.listen_btn.text() == "结束":
            self.listen_btn.setText("试听")

    def on_export(self):
        if not self.parsed_files:
            QMessageBox.warning(self, "提示", "请先选择MIDI文件")
            return
        
        index = self.num_combobox.currentIndex()
        if index < 0:
            QMessageBox.warning(self, "提示", "请选择一个移调方案")
            return
        
        file_name = self.num_combobox.itemData(index)
        
        if not file_name:
            QMessageBox.warning(self, "提示", "无效的文件名")
            return
            
        src_path = os.path.join(self.temp_dir, file_name)
        
        if not os.path.exists(src_path):
            QMessageBox.critical(self, "错误", f"源文件不存在: {src_path}\n请检查文件是否已被删除")
            return
        
        dst_path = os.path.join(self.out_dir, file_name)
        
        counter = 1
        base_name, ext = os.path.splitext(file_name)
        while os.path.exists(dst_path):
            dst_path = os.path.join(self.out_dir, f"{base_name}_{counter}{ext}")
            counter += 1
        
        try:
            shutil.copy2(src_path, dst_path)
            QMessageBox.information(
                self, "导出成功", 
                f"文件已导出到:\n{dst_path}\n\n"
                f"移调: {self.num_combobox.currentText()}\n"
                f"倍速: {self.speed_edit.text() or '1.0'}x"
            )
                
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))


class MidiOptimizer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MIDI优化")
        self.folder = ""
        self.files = []
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)
        layout.setContentsMargins(15, 15, 15, 15)

        # 第一行：文件夹 + 浏览按钮
        row1 = QHBoxLayout()
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("选择文件夹...")
        self.path_edit.setReadOnly(True)
        btn_browse = QPushButton("浏览...")
        btn_browse.setFixedWidth(60)
        btn_browse.clicked.connect(self.browse_folder)
        row1.addWidget(self.path_edit)
        row1.addWidget(btn_browse)
        layout.addLayout(row1)

        # 第二行：文件选择 + 限制数
        row2 = QHBoxLayout()
        row2.addWidget(QLabel("文件:"))
        self.file_combo = QComboBox()
        self.file_combo.setEnabled(False)
        row2.addWidget(self.file_combo, 1)

        row2.addWidget(QLabel("限制:"))
        self.limit_combo = QComboBox()
        self.limit_combo.setFixedWidth(50)
        for i in range(1, 7):
            self.limit_combo.addItem(str(i), i)
        self.limit_combo.setCurrentIndex(5)  # 默认6
        row2.addWidget(self.limit_combo)
        layout.addLayout(row2)

        # 第三行：优化按钮
        self.btn_optimize = QPushButton("🚀 优化")
        self.btn_optimize.setEnabled(False)
        self.btn_optimize.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #1976D2; }
            QPushButton:disabled { background-color: #ccc; }
        """)
        self.btn_optimize.clicked.connect(self.optimize)
        layout.addWidget(self.btn_optimize)

    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.folder = folder
            self.path_edit.setText(folder)
            self.refresh_files()

    def refresh_files(self):
        self.file_combo.clear()
        self.files = [f for f in os.listdir(self.folder) if f.endswith('.txt')]

        if not self.files:
            QMessageBox.warning(self, "提示", "该文件夹没有txt文件")
            self.file_combo.setEnabled(False)
            self.btn_optimize.setEnabled(False)
            return

        for f in self.files:
            self.file_combo.addItem(f, os.path.join(self.folder, f))

        self.file_combo.setEnabled(True)
        self.btn_optimize.setEnabled(True)

    def parse_file(self, filepath):
        events = []
        with open(filepath, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = line.split('\t')
                if len(parts) != 3:
                    continue
                t, k, ts = parts
                if t not in ('P', 'R'):
                    continue
                try:
                    ts = int(ts)
                except:
                    continue
                events.append({'type': t, 'key': k, 'timestamp': ts, 'processed': False})
        return events

    def process(self, events, max_active):
        pitch_dup = release_adv = press_delay = bad = 0
        active = []
        i = 0

        while i < len(events):
            e = events[i]
            if e['processed']:
                i += 1
                continue

            if e['type'] == 'P':
                # 检查同音
                exist_idx = None
                for idx, a in enumerate(active):
                    if a['key'] == e['key']:
                        exist_idx = idx
                        break

                if exist_idx is not None:
                    # 同音处理
                    pitch_dup += 1
                    old = active[exist_idx]
                    # 找对应R
                    for j in range(i+1, len(events)):
                        if events[j]['type'] == 'R' and events[j]['key'] == e['key'] and not events[j]['processed']:
                            events[j]['processed'] = True
                            break
                    # 计算新时间戳
                    if e['timestamp'] - old['timestamp'] > 80:
                        new_ts = e['timestamp'] - 40
                    else:
                        new_ts = (old['timestamp'] + e['timestamp']) // 2
                        bad += 1
                    # 插入R
                    events.insert(i, {'type': 'R', 'key': e['key'], 'timestamp': new_ts, 'processed': False})
                    i += 1
                    active.pop(exist_idx)
                    active.append(e)
                    i += 1
                    continue

                elif len(active) >= max_active:
                    # 超限处理
                    oldest = active[0]
                    for j in range(i+1, len(events)):
                        if events[j]['type'] == 'R' and events[j]['key'] == oldest['key'] and not events[j]['processed']:
                            events[j]['processed'] = True
                            break

                    if e['timestamp'] - oldest['timestamp'] > 40:
                        release_adv += 1
                        new_ts = e['timestamp']
                    else:
                        press_delay += 1
                        new_ts = oldest['timestamp'] + 40
                        e['timestamp'] = new_ts

                    events.insert(i, {'type': 'R', 'key': oldest['key'], 'timestamp': new_ts, 'processed': False})
                    i += 1
                    active.pop(0)
                    active.append(e)
                    i += 1
                    continue
                else:
                    active.append(e)
                    i += 1
                    continue

            elif e['type'] == 'R':
                if not e['processed']:
                    for idx, a in enumerate(active):
                        if a['key'] == e['key']:
                            active.pop(idx)
                            break
                i += 1
                continue

            i += 1

        # 过滤并排序
        result = [e for e in events if not e['processed']]
        result.sort(key=lambda x: (x['timestamp'], 0 if x['type'] == 'R' else 1))
        for e in result:
            e.pop('processed', None)

        return result, pitch_dup, release_adv, press_delay, bad

    def save_result(self, result, orig_path, max_active):
        dir_path = os.path.dirname(os.path.abspath(orig_path))
        basename = os.path.basename(orig_path)
        new_name = f"[M{max_active}]{basename}"
        out_path = os.path.join(dir_path, new_name)

        with open(out_path, 'w', encoding='utf-8') as f:
            for e in result:
                f.write(f"{e['type']}\t{e['key']}\t{e['timestamp']}\n")
        return out_path

    def optimize(self):
        if self.file_combo.currentIndex() < 0:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return

        filepath = self.file_combo.currentData()
        max_active = self.limit_combo.currentData()

        try:
            events = self.parse_file(filepath)
            total = len(events)

            if total == 0:
                QMessageBox.warning(self, "错误", "文件为空或格式错误")
                return

            result, dup, adv, delay, bad = self.process(events, max_active)
            out_path = self.save_result(result, filepath, max_active)

            msg = f"✅ 优化完成！\n\n"
            msg += f"原始事件: {total}\n"
            msg += f"同音重复: {dup}\n"
            msg += f"提前释放: {adv}\n"
            msg += f"推迟按下: {delay}\n"
            msg += f"不理想处理: {bad}\n"
            msg += f"输出事件: {len(result)}\n\n"
            msg += f"保存至: {os.path.basename(out_path)}"

            QMessageBox.information(self, "完成", msg)

        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理失败: {str(e)}")





class PlayThread(QThread):
    """使用QThread替代原生threading，确保与Qt信号机制兼容"""
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, player, delay_ms):
        super().__init__()
        self.player = player
        self.delay_ms = delay_ms
        self._stop_requested = False
    
    def stop(self):
        self._stop_requested = True
        self.player._stop.set()
    
    def run(self):
        try:
            self.player._play_loop(self.delay_ms, self._stop_requested)
        except Exception as e:
            self.error.emit(str(e))
        finally:
            self.finished.emit()


class KeyPlayer:
    def __init__(self):
        self.output = None
        self.is_playing = False
        self._stop = threading.Event()
        self.events = []
        self.thread = None  # 现在是PlayThread(QThread)
        self.speed = 1.0
        self.max_active = 0  # 默认无处理
        self._lock = threading.Lock()
    
    def set_max_active(self, max_active):
        """设置最大同时按键数"""
        self.max_active = max_active
    
    def parse_txt(self, path):
        """解析 TSV 并处理按键限制"""
        # 首先解析原始事件
        raw_events = []
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                parts = line.split('\t')
                if len(parts) >= 3:
                    evt_type, key, t_str = parts[0], parts[1], parts[2]
                    try:
                        t_clean = ''.join(c for c in t_str if c.isdigit())
                        t = int(t_clean)
                        if key and key[0] in ALLOWED_CHARS:
                            raw_events.append({
                                'type': evt_type,
                                'key': key,
                                'timestamp': t,
                                'processed': False
                            })
                    except ValueError:
                        continue
        
        # 处理按键限制
        poly_stats = {
            'total': len(raw_events),
            'pitch_dup': 0,
            'release_adv': 0,
            'press_delay': 0,
            'bad': 0,
            'output': 0
        }
        
        if self.max_active > 0:  # 只有当限制大于0时才处理
            processed_events, pitch_dup, release_adv, press_delay, bad = self.process_events(raw_events, self.max_active)
            poly_stats['pitch_dup'] = pitch_dup
            poly_stats['release_adv'] = release_adv
            poly_stats['press_delay'] = press_delay
            poly_stats['bad'] = bad
            poly_stats['output'] = len(processed_events)
        else:
            processed_events = raw_events
            poly_stats['output'] = len(processed_events)
        
        # 转换为播放格式
        events = []
        for e in processed_events:
            if e['type'] in ('P', 'R'):
                events.append({
                    'is_press': e['type'] == 'P',
                    'key': e['key'][0].lower() if e['key'][0].isalpha() else e['key'][0],
                    'time_ms': e['timestamp']
                })
        
        events.sort(key=lambda x: x['time_ms'])
        return events, poly_stats
    
    def process_events(self, events, max_active):
        """处理事件，限制同时按键数"""
        pitch_dup = release_adv = press_delay = bad = 0
        active = []
        i = 0

        while i < len(events):
            e = events[i]
            if e['processed']:
                i += 1
                continue

            if e['type'] == 'P':
                # 检查同音
                exist_idx = None
                for idx, a in enumerate(active):
                    if a['key'] == e['key']:
                        exist_idx = idx
                        break

                if exist_idx is not None:
                    # 同音处理
                    pitch_dup += 1
                    old = active[exist_idx]
                    # 找对应R
                    for j in range(i+1, len(events)):
                        if events[j]['type'] == 'R' and events[j]['key'] == e['key'] and not events[j]['processed']:
                            events[j]['processed'] = True
                            break
                    # 计算新时间戳
                    if e['timestamp'] - old['timestamp'] > 80:
                        new_ts = e['timestamp'] - 40
                    else:
                        new_ts = (old['timestamp'] + e['timestamp']) // 2
                        bad += 1
                    # 插入R
                    events.insert(i, {'type': 'R', 'key': e['key'], 'timestamp': new_ts, 'processed': False})
                    i += 1
                    active.pop(exist_idx)
                    active.append(e)
                    i += 1
                    continue

                elif len(active) >= max_active:
                    # 超限处理
                    oldest = active[0]
                    for j in range(i+1, len(events)):
                        if events[j]['type'] == 'R' and events[j]['key'] == oldest['key'] and not events[j]['processed']:
                            events[j]['processed'] = True
                            break

                    if e['timestamp'] - oldest['timestamp'] > 40:
                        release_adv += 1
                        new_ts = e['timestamp']
                    else:
                        press_delay += 1
                        new_ts = oldest['timestamp'] + 40
                        e['timestamp'] = new_ts

                    events.insert(i, {'type': 'R', 'key': oldest['key'], 'timestamp': new_ts, 'processed': False})
                    i += 1
                    active.pop(0)
                    active.append(e)
                    i += 1
                    continue
                else:
                    active.append(e)
                    i += 1
                    continue

            elif e['type'] == 'R':
                if not e['processed']:
                    for idx, a in enumerate(active):
                        if a['key'] == e['key']:
                            active.pop(idx)
                            break
                i += 1
                continue

            i += 1

        # 过滤并排序
        result = [e for e in events if not e['processed']]
        result.sort(key=lambda x: (x['timestamp'], 0 if x['type'] == 'R' else 1))
        for e in result:
            e.pop('processed', None)

        return result, pitch_dup, release_adv, press_delay, bad
    
    def _play_loop(self, delay_ms: int, stop_requested_flag):
        """播放循环 - 使用超时机制确保可中断"""
        try:
            self.is_playing = True
            
            # 延迟阶段 - 使用小片段检查停止信号
            if delay_ms > 0:
                chunk = 50  # 50ms检查一次
                elapsed = 0
                while elapsed < delay_ms and not self._stop.is_set():
                    remaining = min(chunk, delay_ms - elapsed)
                    time.sleep(remaining / 1000.0)
                    elapsed += remaining
            
            if self._stop.is_set() or not self.events:
                return
            
            start_ns = time.perf_counter_ns()
            idx = 0
            total = len(self.events)
            speed_factor = 1.0 / self.speed
            
            while idx < total and not self._stop.is_set():
                evt = self.events[idx]
                target_ns = start_ns + int(evt['time_ms'] * speed_factor * 1_000_000)
                current_ns = time.perf_counter_ns()
                
                if current_ns >= target_ns:
                    if not self._stop.is_set():
                        if evt['is_press']:
                            if self.output:
                                self.output.press(evt['key'])
                        else:
                            if self.output:
                                self.output.release(evt['key'])
                    idx += 1
                else:
                    # 计算等待时间，但设置最大等待上限确保及时响应停止
                    remaining_ns = target_ns - current_ns
                    wait_ms = remaining_ns / 1_000_000
                    
                    # 最多等待50ms就检查一次停止信号
                    if wait_ms > 50:
                        time.sleep(0.05)
                    elif wait_ms > 5:
                        time.sleep(wait_ms / 1000.0)
                    else:
                        # 小于5ms时忙等待确保精度
                        pass
                    
        except Exception as e:
            print(f"[KeyPlayer] 播放循环错误: {e}")
            raise
        finally:
            # 确保释放所有按键
            try:
                if self.output:
                    self.output.release_all()
            except Exception as e:
                print(f"[KeyPlayer] 释放按键失败: {e}")
            
            with self._lock:
                self.is_playing = False
    
    def play(self, path: str, speed: float = 1.0, delay_ms: int = 2000):
        with self._lock:
            if self.is_playing:
                return False, None
            
            self.events, poly_stats = self.parse_txt(path)
            if not self.events:
                return False, None
            
            self.speed = speed
            self._stop.clear()
            self.is_playing = True  # 先设置状态再启动线程
        
        # 使用QThread确保与Qt主线程兼容
        self.thread = PlayThread(self, delay_ms)
        self.thread.start()
        return True, poly_stats
    
    def stop(self):
        """停止播放 - 确保能中断延迟等待"""
        self._stop.set()
        
        # 如果线程存在，等待它结束（带超时）
        if self.thread:
            # 使用QThread的quit/wait机制
            if self.thread.isRunning():
                # 最多等待1秒，强制结束
                self.thread.wait(1000)
                if self.thread.isRunning():
                    print("[KeyPlayer] 线程未响应，强制终止")
                    self.thread.terminate()
                    self.thread.wait(500)
        
        # 确保状态重置
        with self._lock:
            self.is_playing = False
        
        # 安全释放按键
        try:
            if self.output:
                self.output.release_all()
        except Exception as e:
            print(f"[KeyPlayer] stop中释放失败: {e}")


class AutoKeyWidget(QWidget):
    arduino_changed = pyqtSignal(bool, str)
    
    def __init__(self):
        super().__init__()
        self.player = KeyPlayer()
        self.arduino = None
        self.virtual = None
        self.serial = None
        self.current_folder = ""
        self.port_name = ""
        self.output_mode = "arduino"
        self._build_ui()
        self._init_timers()
        self._scan_arduino()
        self._update_status_text()
    
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(6)
        layout.setContentsMargins(10, 8, 10, 8)
        
        row1 = QHBoxLayout()
        row1.setSpacing(6)
        
        self.folder_btn = QPushButton("📁")
        self.folder_btn.setFixedSize(30, 26)
        self.folder_btn.setToolTip("选择乐谱文件夹")
        self.folder_btn.clicked.connect(self._select_folder)
        
        self.file_combo = QComboBox()
        self.file_combo.setFixedHeight(26)
        self.file_combo.setMinimumWidth(150)
        self.file_combo.currentTextChanged.connect(self._on_file_changed)
        
        output_label = QLabel("输出:")
        output_label.setFixedHeight(26)
        row1.addWidget(output_label)
        
        self.output_combo = QComboBox()
        self.output_combo.setFixedWidth(100)
        self.output_combo.setFixedHeight(26)
        self.output_combo.addItem("Arduino", "arduino")
        self.output_combo.addItem("虚拟按键", "virtual")
        self.output_combo.setCurrentIndex(0)
        self.output_combo.currentIndexChanged.connect(self._on_output_changed)
        row1.addWidget(self.output_combo)
        
        self.status_lbl = QLabel("未连接")
        self.status_lbl.setStyleSheet("color: #e74c3c; font-size: 11px; font-weight: bold;")
        self.status_lbl.setFixedHeight(26)
        self.status_lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        
        row1.addWidget(self.folder_btn)
        row1.addWidget(self.file_combo, 1)
        row1.addWidget(self.status_lbl)
        
        layout.addLayout(row1)
        
        row2 = QHBoxLayout()
        row2.setSpacing(6)
        
        speed_label = QLabel("倍速:")
        speed_label.setFixedHeight(26)
        row2.addWidget(speed_label)
        
        self.speed_input = QLineEdit("1.0")
        self.speed_input.setFixedSize(55, 26)
        self.speed_input.setMinimumSize(55, 26)
        self.speed_input.setValidator(QDoubleValidator(0.1, 5.0, 2))
        self.speed_input.setAlignment(Qt.AlignCenter)
        self.speed_input.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        row2.addWidget(self.speed_input)
        
        delay_label = QLabel("延时:")
        delay_label.setFixedHeight(26)
        row2.addWidget(delay_label)
        
        self.delay_input = QLineEdit("2000")
        self.delay_input.setFixedSize(60, 26)
        self.delay_input.setMinimumSize(60, 26)
        self.delay_input.setValidator(QIntValidator(0, 30000))
        self.delay_input.setAlignment(Qt.AlignCenter)
        self.delay_input.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        row2.addWidget(self.delay_input)
        
        ms_label = QLabel("ms")
        ms_label.setFixedHeight(26)
        row2.addWidget(ms_label)
        
        # 添加最大按键数下拉框
        poly_label = QLabel("按键数:")
        poly_label.setFixedHeight(26)
        row2.addWidget(poly_label)
        
        self.poly_combo = QComboBox()
        self.poly_combo.setFixedWidth(100)
        self.poly_combo.addItem("无处理（默认）", 0)
        for i in range(1, 7):
            self.poly_combo.addItem(str(i), i)
        self.poly_combo.setCurrentIndex(0)  # 默认无处理
        self.poly_combo.currentIndexChanged.connect(self._on_poly_changed)
        row2.addWidget(self.poly_combo)
        
        row2.addStretch(1)
        
        self.play_btn = QPushButton("▶ 播放")
        self.play_btn.setFixedSize(75, 28)
        self.play_btn.setEnabled(False)
        self.play_btn.setStyleSheet("""
            QPushButton { 
                background-color: #95a5a6; 
                color: white; 
                border-radius: 4px; 
                font-weight: bold;
                font-size: 12px;
                padding: 2px;
            }
            QPushButton:enabled { background-color: #27ae60; }
            QPushButton:enabled:hover { background-color: #2ecc71; }
        """)
        self.play_btn.clicked.connect(self._toggle_play)
        
        row2.addWidget(self.play_btn)
        
        layout.addLayout(row2)
        
        # 添加Poly处理结果文本框
        self.poly_result_text = QLabel("Poly处理结果将显示在这里")
        self.poly_result_text.setStyleSheet("""
            QLabel {
                border: 1px solid #ddd;
                border-radius: 4px;
                padding: 8px;
                background-color: #f9f9f9;
                font-family: Consolas, monospace;
                font-size: 10px;
                color: #333;
                min-height: 60px;
            }
        """)
        self.poly_result_text.setWordWrap(True)
        layout.addWidget(self.poly_result_text)
        
        self.arduino_changed.connect(self._update_status)
    
    def _on_file_changed(self, text):
        self._update_play_button_state()
        self._update_status_text()
    

    
    def _on_output_changed(self, index):
        """处理输出方式变更"""
        if self.player.is_playing:
            self.player.stop()
            self._update_play_btn(False)
        
        new_mode = self.output_combo.currentData()
        self.output_mode = new_mode
        print(f"[AutoKey] 输出方式切换为: {new_mode}")
        
        self._disconnect()
        
        if new_mode == "virtual":
            self.virtual = VirtualOutput(debug=True)
            self.player.output = self.virtual
            self.arduino_changed.emit(True, "虚拟按键")
            self._update_play_button_state()
        else:
            self._scan_arduino(force=True)
        
        self._update_status_text()
    
    def _on_poly_changed(self, index):
        """处理最大按键数变更"""
        max_active = self.poly_combo.currentData()
        self.player.set_max_active(max_active)
        print(f"[AutoKey] 最大按键数设置为: {max_active}")
        
        self._update_status_text()
    
    def _update_status_text(self):
        """更新状态文本框显示"""
        file_text = ""
        if self.file_combo.currentText():
            file_text = f"当前选择文件: {self.file_combo.currentText()}\n\n"
        
        status_text = ""
        if self.output_mode == "virtual":
            status_text = "输出方式: 虚拟按键 (已就绪)\n\n"
        else:
            if self.arduino:
                status_text = f"输出方式: Arduino ({self.status_lbl.text()})\n\n"
            else:
                status_text = "输出方式: Arduino (未连接)\n\n"
        
        max_active = self.player.max_active
        poly_text = f"当前按键数限制: {max_active}"
        if max_active == 0:
            poly_text += " (无处理)"
        
        self.poly_result_text.setText(file_text + status_text + poly_text)
    
    def _init_timers(self):
        self.file_timer = QTimer()
        self.file_timer.timeout.connect(self._refresh_files)
        self.file_timer.start(2000)

        self.arduino_timer = QTimer()
        self.arduino_timer.timeout.connect(self._heartbeat)
        self.arduino_timer.start(2000)

        self.reconnect_timer = QTimer()
        self.reconnect_timer.timeout.connect(self._try_reconnect)
        self.reconnect_timer.start(3000)
        
        # 新增：UI刷新定时器，确保播放按钮状态与实际一致
        self.ui_timer = QTimer()
        self.ui_timer.timeout.connect(self._sync_ui_state)
        self.ui_timer.start(100)  # 100ms检查一次
    
    def _sync_ui_state(self):
        """同步UI状态 - 确保播放按钮与实际播放状态一致"""
        actual_playing = self.player.is_playing
        btn_showing_stop = self.play_btn.text() == "■ 停止"
        
        # 如果状态不一致，强制同步
        if actual_playing != btn_showing_stop:
            print(f"[UI] 状态同步: is_playing={actual_playing}, btn={btn_showing_stop}")
            self._update_play_btn(actual_playing)
            self._update_play_button_state()
    
    def _select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择乐谱文件夹")
        if folder:
            self.current_folder = folder
            self._refresh_files()
    
    def _refresh_files(self):
        if not self.current_folder:
            return
        
        current = self.file_combo.currentText()
        files = sorted(glob.glob(os.path.join(self.current_folder, "*.txt")))
        
        # 过滤掉带 [Mx] 标记的文件
        filtered_files = []
        for f in files:
            basename = os.path.basename(f)
            if not basename.startswith('[M'):
                filtered_files.append(f)
        
        basenames = [os.path.basename(f) for f in filtered_files]
        
        current_items = [self.file_combo.itemText(i) for i in range(self.file_combo.count())]
        if basenames != current_items:
            self.file_combo.clear()
            self.file_combo.addItems(basenames)
            if current in basenames:
                self.file_combo.setCurrentText(current)
        
        self._update_play_button_state()
    
    def _update_play_button_state(self):
        has_file = bool(self.file_combo.currentText())
        has_output = False
        if self.output_mode == "virtual":
            has_output = self.virtual is not None
        else:
            has_output = self.arduino is not None
        # 播放中也可以点击停止
        can_click = (has_file and has_output) or self.player.is_playing
        self.play_btn.setEnabled(can_click)
    
    def _scan_arduino(self, force=False):
        if force and self.arduino:
            self._disconnect()
        
        if self.arduino and not force:
            return
        
        ports = [
            p for p in serial.tools.list_ports.comports()
            if 'Arduino' in p.description or 'CH340' in p.description
            or p.vid in (0x2341, 0x1A86, 0x2A03, 0x10C4, 0x0403)
        ]
        
        if not ports:
            return
        
        for port in ports:
            try:
                ser = serial.Serial(port.device, 115200, timeout=1)
                time.sleep(1.5)
                
                ser.reset_input_buffer()
                ser.write(b"I\n")
                ser.flush()
                
                start = time.time()
                while time.time() - start < 2.0:
                    if ser.in_waiting:
                        resp = ser.readline().decode().strip()
                        if resp == "R":
                            self.serial = ser
                            self.arduino = ArduinoOutput(ser)
                            self.player.output = self.arduino
                            self.port_name = port.device.split('/')[-1]
                            if self.port_name.startswith('tty'):
                                self.port_name = port.device
                            
                            self.arduino_changed.emit(True, self.port_name)
                            self._update_play_button_state()
                            
                            print(f"[AutoKey] Arduino已连接: {self.port_name}")
                            return
                    time.sleep(0.05)
                
                ser.close()
            except Exception as e:
                print(f"[AutoKey] 连接失败 {port.device}: {e}")
                try:
                    ser.close()
                except:
                    pass
    
    def _heartbeat(self):
        if self.output_mode != "arduino":
            return
        if not self.serial or not self.serial.is_open:
            if self.arduino:
                self._disconnect()
            return

        try:
            self.serial.write(b"P\n")
        except Exception as e:
            print(f"[AutoKey] 心跳失败: {e}")
            self._disconnect()
    
    def _disconnect(self):
        was_connected = (self.arduino is not None) or (self.virtual is not None)
        if was_connected:
            print("[AutoKey] 输出断开")
        
        # 关键：先停止播放
        if self.player.is_playing:
            try:
                self.player.stop()
                self._update_play_btn(False)
            except Exception as e:
                print(f"[AutoKey] 停止播放失败: {e}")
        
        if self.arduino:
            self.arduino._alive = False
        if self.virtual:
            self.virtual._alive = False
        
        self.player.output = None

        if self.serial:
            try:
                try:
                    self.serial.write(b"X\n")
                    time.sleep(0.05)
                except:
                    pass
                self.serial.close()
            except Exception as e:
                print(f"[AutoKey] 关闭串口失败: {e}")

        self.serial = None
        self.arduino = None
        self.virtual = None
        self.port_name = ""

        if was_connected:
            self.arduino_changed.emit(False, "未连接")
            self._update_play_button_state()

    def _try_reconnect(self):
        if self.output_mode != "arduino":
            return
        if not self.arduino and not self.player.is_playing:
            print("[AutoKey] 尝试重连...")
            self._scan_arduino(force=True)

    def _update_status(self, connected, text):
        self.status_lbl.setText(text)
        self.status_lbl.setStyleSheet(
            "color: #27ae60; font-size: 11px; font-weight: bold;" if connected else "color: #e74c3c; font-size: 11px; font-weight: bold;"
        )
        
        # 更新结果文本框显示连接状态
        current_text = self.poly_result_text.text()
        if not current_text.startswith("Arduino") and not current_text.startswith("Poly处理结果"):
            if connected:
                status_text = f"Arduino连接状态: 已连接 ({text})\n\n"
                status_text += "Poly处理结果将在播放时显示"
                self.poly_result_text.setText(status_text)
            else:
                status_text = f"Arduino连接状态: 未连接\n\n"
                status_text += "请连接Arduino设备后再播放"
                self.poly_result_text.setText(status_text)
    
    def _update_play_btn(self, playing):
        self.play_btn.setText("■ 停止" if playing else "▶ 播放")
        self.play_btn.setStyleSheet("""
            QPushButton { 
                background-color: #e74c3c; 
                color: white; 
                border-radius: 4px; 
                font-weight: bold;
                font-size: 12px;
                padding: 2px;
            }
            QPushButton:hover { background-color: #c0392b; }
        """ if playing else """
            QPushButton { 
                background-color: #27ae60; 
                color: white; 
                border-radius: 4px; 
                font-weight: bold;
                font-size: 12px;
                padding: 2px;
            }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        self._update_play_button_state()
    
    def _toggle_play(self):
        if self.player.is_playing:
            print("[AutoKey] 用户点击停止")
            self.player.stop()
            self._update_play_btn(False)
            
            # 更新结果文本框显示停止信息
            file_text = ""
            if self.file_combo.currentText():
                file_text = f"当前选择文件: {self.file_combo.currentText()}\n\n"
            
            status_text = ""
            if self.arduino:
                status_text = f"Arduino连接状态: 已连接 ({self.status_lbl.text()})\n\n"
            else:
                status_text = "Arduino连接状态: 未连接\n\n"
            
            play_text = "播放已停止\n\n"
            poly_text = f"当前按键数限制: {self.player.max_active}"
            if self.player.max_active == 0:
                poly_text += " (无处理)"
            
            self.poly_result_text.setText(file_text + status_text + play_text + poly_text)
        else:
            if not self.file_combo.currentText():
                return
            has_output = False
            if self.output_mode == "virtual":
                has_output = self.virtual is not None
            else:
                has_output = self.arduino is not None
            if not has_output:
                return
            
            path = os.path.join(self.current_folder, self.file_combo.currentText())
            
            try:
                speed = float(self.speed_input.text() or "1.0")
                delay = int(self.delay_input.text() or "2000")
            except:
                speed = 1.0
                delay = 2000
            
            success, poly_stats = self.player.play(path, speed, delay)
            if success:
                self._update_play_btn(True)
                # 更新Poly处理结果
                if poly_stats:
                    max_active = self.player.max_active
                    if max_active > 0:
                        result_text = f"Poly处理结果:\n"
                        result_text += f"原始事件: {poly_stats['total']}\n"
                        result_text += f"同音重复: {poly_stats['pitch_dup']}\n"
                        result_text += f"提前释放: {poly_stats['release_adv']}\n"
                        result_text += f"推迟按下: {poly_stats['press_delay']}\n"
                        result_text += f"不理想处理: {poly_stats['bad']}\n"
                        result_text += f"输出事件: {poly_stats['output']}\n"
                        result_text += f"限制按键数: {max_active}\n\n"
                        
                        # 生成处理后的文件
                        base_dir = get_base_dir()
                        out_dir = os.path.join(base_dir, "out")
                        os.makedirs(out_dir, exist_ok=True)
                        
                        file_name = os.path.basename(path)
                        base_name, ext = os.path.splitext(file_name)
                        processed_file_name = f"[M{max_active}]{base_name}{ext}"
                        processed_file_path = os.path.join(out_dir, processed_file_name)
                        
                        # 保存处理后的事件
                        with open(processed_file_path, 'w', encoding='utf-8') as f:
                            for event in self.player.events:
                                evt_type = 'P' if event['is_press'] else 'R'
                                f.write(f"{evt_type}\t{event['key']}\t{event['time_ms']}\n")
                        
                        result_text += f"已生成处理文件: {processed_file_name}\n"
                        result_text += f"保存位置: {out_dir}\n\n"
                        result_text += f"正在播放..."
                    else:
                        result_text = f"Poly处理结果:\n"
                        result_text += f"原始事件: {poly_stats['total']}\n"
                        result_text += f"输出事件: {poly_stats['output']}\n"
                        result_text += f"限制按键数: 无处理\n\n"
                        result_text += f"正在播放..."
                    self.poly_result_text.setText(result_text)
                else:
                    self.poly_result_text.setText("Poly处理结果将显示在这里")
    
    def closeEvent(self, event):
        self.player.stop()
        if self.serial:
            try:
                self.serial.write(b"X\n")
                time.sleep(0.1)
                self.serial.close()
            except:
                pass
        event.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoMid")
        self.setFixedSize(680, 320)
        
        # 创建主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # 顶置功能复选框和状态显示
        top_layout = QHBoxLayout()
        self.always_on_top_checkbox = QCheckBox("窗口顶置")
        self.always_on_top_checkbox.setFont(QFont("微软雅黑", 9))
        self.always_on_top_checkbox.stateChanged.connect(self.toggle_always_on_top)
        top_layout.addWidget(self.always_on_top_checkbox)
        
        # 状态显示标签
        self.status_label = QLabel("就绪")
        self.status_label.setFont(QFont("微软雅黑", 9))
        self.status_label.setStyleSheet("color: #27ae60;")
        top_layout.addStretch(1)
        top_layout.addWidget(self.status_label)
        main_layout.addLayout(top_layout)
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)
        
        # 创建各个功能选项卡
        self.mid_translator_tab = AutoMidWindow()
        self.auto_key_tab = AutoKeyWidget()
        
        # 添加选项卡
        self.tab_widget.addTab(self.mid_translator_tab, "MIDI转换")
        self.tab_widget.addTab(self.auto_key_tab, "AutoKey")
    
    def toggle_always_on_top(self, state):
        """切换窗口顶置状态"""
        if state == Qt.Checked:
            self.setWindowFlags(self.windowFlags() | Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~Qt.WindowStaysOnTopHint)
        self.show()


if __name__ == "__main__":
    from PyQt5.QtCore import QCoreApplication
    QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    app = QApplication(sys.argv)
    font = QFont("微软雅黑", 9)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())