#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diablo 4 宏程序 - 主程序
"""

import time
import threading
import xml.etree.ElementTree as ET
import keyboard
import mouse
import re
import sys
import os
from collections import deque
from typing import Dict, List, Optional, Callable, Any

# 彩色输出支持
try:
    from colorama import Fore, Back, Style, init as colorama_init
    colorama_init(autoreset=True)
    COLORAMA_AVAILABLE = True
except ImportError:
    COLORAMA_AVAILABLE = False
    class _DummyColor:
        def __getattr__(self, name): return ''
    Fore = Back = Style = _DummyColor()

# Windows 窗口检测
try:
    import win32gui
    WINDOW_DETECTION_AVAILABLE = True
except ImportError:
    WINDOW_DETECTION_AVAILABLE = False


# 程序配置常量
class Config:
    """程序配置"""
    APP_NAME = "Diablo 4 Macro"
    APP_VERSION = "1.0.0"
    CONFIG_FILE = "macro_config.txt"
    LAST_UPDATE = "2025-10-22"

    # 目标窗口配置
    TARGET_WINDOWS = ['Diablo IV', '暗黑破坏神IV']
    # TARGET_WINDOWS = []  # 留空或设置为 None 表示不限制窗口

    # 控制热键
    HOTKEY_PAUSE = 'F10'
    HOTKEY_RELOAD = 'F11'
    HOTKEY_EXIT = 'F12'

    # 按键输入缓冲配置
    INPUT_BUFFER_ENABLED = True        # 启用按键输入缓冲
    INPUT_BUFFER_SIZE = 3              # 最大缓冲队列长度
    RAPID_PRESS_THRESHOLD = 0.2        # 连按时间阈值（秒），在此时间内的重复按键才会缓冲
  
    # 时间间隔常量（秒）  
    KEY_PRESS_INTERVAL = 0.01          # 按键后等待时间
    DELAY_CHECK_INTERVAL = 0.05        # 延迟可中断检查间隔
    DEFAULT_DELAY_FALLBACK = 0.1       # 延迟解析失败时的默认值
    ESC_DOUBLE_CLICK_INTERVAL = 0.5    # Esc 双击间隔
    CONTROL_KEY_DEBOUNCE = 0.1         # 控制热键防抖间隔
    THREAD_JOIN_TIMEOUT = 1.0          # 线程join超时时间
    WINDOW_FOCUS_CHECK_INTERVAL = 0.1  # 窗口焦点检查间隔

    # 安全保护
    FORCE_RELEASE_MODIFIER_KEYS = ['shift', 'ctrl', 'alt', 'win']  # 强制释放的修饰键

    # UI 显示
    SEPARATOR_LINE = "=" * 50


# ========================================
# 语言映射表
# ========================================
class LanguageMapping:

    # 英文同义词到标准英文的映射
    EN_SYNONYMS = {
        'yes': 'true',
        'no': 'false',
        'wait': 'delay',
        'control': 'ctrl',
        'trigger': 'trigger_key',
    }

    # 统一的中文到英文映射表
    ZH_CN = {
        # 配置项
        '名称': 'name', '说明': 'description',
        '触发键': 'trigger_key', '循环': 'loop', '重复': 'repeat', '动作': 'actions',
        '重置按键': 'reset_keys', '附加按键': 'additional_keys', '默认延迟': 'default_delay',
        '跳过窗口检测': 'skip_window_check', '忽略窗口检测': 'skip_window_check',
        '结束按键': 'finish_actions', '完成按键': 'finish_actions',
        '结束动作': 'finish_actions', '完成动作': 'finish_actions',
        '起始动作': 'start_actions', '开始动作': 'start_actions',
        # 重复模式
        '按住时': 'hold',
        # 动作指令
        '按下': 'press', '按住': 'keydown', '松开': 'keyup', '释放': 'keyup',
        '等待': 'delay', '延迟': 'delay',
        # 布尔值 (true)
        '是': 'true', '打开': 'true', '启用': 'true', '持续': 'true',
        # 布尔值 (false)
        '否': 'false', '关闭': 'false', '禁用': 'false', '停止': 'false',
        # 特殊键
        '空格': 'space', '回车': 'enter', '退格': 'backspace', '删除': 'delete',
        '上': 'up', '下': 'down', '左': 'left', '右': 'right',
        # 鼠标键（单击）
        '左键': 'leftclick', '右键': 'rightclick', '中键': 'middleclick',
        '侧键1': 'x1click', '侧键2': 'x2click',
        # 鼠标键（双击）
        '双击左键': 'leftdoubleclick', '双击右键': 'rightdoubleclick',
        '双击中键': 'middledoubleclick', '双击侧键1': 'x1doubleclick',
        '双击侧键2': 'x2doubleclick',
        # 小键盘
        '小键盘0': 'keypad 0', '小键盘1': 'keypad 1', '小键盘2': 'keypad 2',
        '小键盘3': 'keypad 3', '小键盘4': 'keypad 4', '小键盘5': 'keypad 5',
        '小键盘6': 'keypad 6', '小键盘7': 'keypad 7', '小键盘8': 'keypad 8',
        '小键盘9': 'keypad 9', '小键盘点': 'keypad .',
        '小键盘加': 'keypad +', '小键盘减': 'keypad -',
        '小键盘乘': 'keypad *', '小键盘除': 'keypad /',
        '小键盘回车': 'keypad enter',
    }

    # XML 格式的重复模式映射（数字 -> 英文）
    XML_REPEAT_MODE = {'0': 'once', '1': 'loop', '2': 'hold'}

    # 鼠标动作映射（用于动作解析）
    MOUSE_ACTIONS = {
        # 单击
        'leftclick': lambda _: {'type': 'click', 'button': 'left'},
        'rightclick': lambda _: {'type': 'click', 'button': 'right'},
        'middleclick': lambda _: {'type': 'click', 'button': 'middle'},
        'x1click': lambda _: {'type': 'click', 'button': 'x1'},
        'x2click': lambda _: {'type': 'click', 'button': 'x2'},
        # 双击
        'leftdoubleclick': lambda _: {'type': 'doubleclick', 'button': 'left'},
        'rightdoubleclick': lambda _: {'type': 'doubleclick', 'button': 'right'},
        'middledoubleclick': lambda _: {'type': 'doubleclick', 'button': 'middle'},
        'x1doubleclick': lambda _: {'type': 'doubleclick', 'button': 'x1'},
        'x2doubleclick': lambda _: {'type': 'doubleclick', 'button': 'x2'},
    }

    # UI 消息映射
    UI_MESSAGES = {
        # 警告和错误信息
        'warn_no_pywin32': '警告: 未安装 pywin32，窗口焦点检测功能不可用',
        'warn_need_pywin32': '[!] 窗口焦点检测需要 pywin32 库，请运行: pip install pywin32',
        'warn_window_lost': '[!] 窗口焦点丢失，自动停止宏执行',
        'warn_config_format': '[!] 警告: 配置行格式错误（缺少\'=\'）: {0}',
        'warn_unknown_config': '[!] 警告: 未知的配置项 \'{0}\' 在宏 \'{1}\'',
        'warn_parse_failed': '[!] 警告: 解析配置行失败: {0}',
        'warn_error': '   错误: {0}',
        'warn_duplicate_key': '[!] 警告: 触发键 \'{0}\' 重复!',
        'warn_key_override': '   宏 \'{0}\' 将被覆盖为 \'{1}\'',
        'warn_buffer_full': '[!] 按键缓冲队列已满 ({0})，跳过宏: {1}',

        # 解析和加载信息
        'unnamed_macro': '未命名宏',
        'parse_xml_failed': '解析 XML 格式失败: {0}',
        'load_file_failed': '加载文件失败: {0}',
        'loading_config': '正在加载配置文件: {0}',
        'no_macros_found': '未找到有效的宏配置',
        'macros_loaded': '\n成功加载 {0} 个宏:',
        'macro_info': '  [{0}] {1}',
        'macro_details': '      触发键: {0} | 模式: {1} | 动作数: {2}',
        'macro_description': '      说明: {0}',

        # 宏模式名称
        'mode_once': '单次',
        'mode_loop': '循环',
        'mode_hold': '按住',
        'trigger_not_set': '未设置',

        # 状态信息
        'status_paused': '宏监听已暂停',
        'status_resume_hint': '按 {0} 恢复监听',
        'status_resumed': '宏监听已恢复',
        'status_reloading': '正在重新加载配置...',
        'status_reloaded': '配置已重新加载！\n',
        'status_buffer_add': '宏已加入缓冲队列: {0} (队列长度: {1})',
        'status_interrupt': '中断当前宏，立即执行: {0}',
        'status_buffer_exec': '从缓冲队列执行: {0} (剩余: {1})',
        'status_buffer_clear': '已清空缓冲队列 ({0} 个宏)',
        'status_emergency_stop': '紧急停止！所有宏已停止',
        'status_keys_reset': '所有按键已重置',

        # 启动信息
        'app_started': '宏程序已启动！',
        'hint_press_hotkey': '按下配置的热键来执行宏',
        'hint_pause': '按 {0} 暂停/恢复监听',
        'hint_reload': '按 {0} 重新加载配置',
        'hint_exit': '按 {0} 退出程序',
        'hint_emergency': '双击 ESC 紧急停止所有宏',
        'buffer_enabled': '[OK] 按键输入缓冲已启用 (队列: {0})',

        # 控制操作
        'pause_on': '已暂停',
        'pause_off': '已恢复',
        'config_reloaded': '配置已重载',
        'emergency_stopped': '急停',
        'app_start': '宏程序已启动',
        'app_hints': '[{0} 暂停 | {1} 重载 | {2} 退出 | ESC×2 急停]',
        'app_exiting': '退出...',
        'app_stopped': '程序已停止',

        # 配置文件相关
        'config_not_found': '配置文件不存在: {0}',
        'config_ensure_exists': '请确保 {0} 存在于当前目录',
        'press_enter_exit': '\n按回车键退出...',
        'using_config': '\n使用配置文件: {0}',

        # 依赖警告
        'pywin32_not_installed': 'pywin32 未安装，窗口检测不可用',

        # 窗口焦点安全锁
        'window_lock': '窗口焦点安全锁: {0}',
        'window_lock_enabled': '已启用',
        'window_lock_disabled': '未启用（需要 pywin32）',
        'target_windows': '目标窗口: {0}',
        'hint_auto_stop': '提示: 切换窗口将自动停止所有宏\n',
    }

    @staticmethod
    def msg(key: str, *args) -> str:
        """获取格式化的 UI 消息"""
        template = LanguageMapping.UI_MESSAGES.get(key, key)
        if args:
            return template.format(*args)
        return template


# 简短的辅助引用
_msg = LanguageMapping.msg


# ========================================
# 彩色日志系统
# ========================================
class Logger:
    """彩色日志输出系统"""

    # 日志级别颜色配置
    COLORS = {
        'success': Fore.GREEN,      # 成功 - 绿色
        'info': Fore.CYAN,          # 信息 - 青色
        'warning': Fore.YELLOW,     # 警告 - 黄色
        'error': Fore.RED,          # 错误 - 红色
        'status': Fore.MAGENTA,     # 状态 - 洋红色
        'macro': Fore.BLUE,         # 宏操作 - 蓝色
        'dim': Style.DIM,           # 暗淡 - 次要信息
        'bright': Style.BRIGHT,     # 明亮 - 重要信息
    }

    # 日志图标
    ICONS = {
        'success': '[OK]',
        'info': '[*]',
        'warning': '[!]',
        'error': '[X]',
        'status': '[+]',
        'macro': '[>]',
        'buffer': '[..]',
        'stop': '[#]',
    }

    @staticmethod
    def _format_message(level: str, message: str, use_icon: bool = True) -> str:
        """格式化消息"""
        color = Logger.COLORS.get(level, '')
        icon = Logger.ICONS.get(level, '') if use_icon else ''
        prefix = f"{icon} " if icon else ""
        return f"{color}{prefix}{message}{Style.RESET_ALL}"

    @staticmethod
    def success(message: str, use_icon: bool = True):
        """成功消息 - 绿色"""
        print(Logger._format_message('success', message, use_icon))

    @staticmethod
    def info(message: str, use_icon: bool = True):
        """信息消息 - 青色"""
        print(Logger._format_message('info', message, use_icon))

    @staticmethod
    def warning(message: str, use_icon: bool = True):
        """警告消息 - 黄色"""
        print(Logger._format_message('warning', message, use_icon))

    @staticmethod
    def error(message: str, use_icon: bool = True):
        """错误消息 - 红色"""
        print(Logger._format_message('error', message, use_icon))

    @staticmethod
    def status(message: str, use_icon: bool = True):
        """状态消息 - 洋红色"""
        print(Logger._format_message('status', message, use_icon))

    @staticmethod
    def macro(message: str, use_icon: bool = True):
        """宏操作消息 - 蓝色"""
        print(Logger._format_message('macro', message, use_icon))


# 简短的日志引用
log = Logger


class WindowMonitor:
    """窗口焦点监控器 - 用于检测游戏窗口是否活动"""

    def __init__(self, target_window_names: Optional[List[str]] = None):
        """
        初始化窗口监控器
        """
        self.target_window_names = target_window_names
        # 只有当提供了非空列表时才启用窗口检测
        self.enabled = (WINDOW_DETECTION_AVAILABLE and
                       target_window_names is not None and
                       len(target_window_names) > 0)

        if not WINDOW_DETECTION_AVAILABLE and target_window_names:
            log.warning(_msg('warn_need_pywin32'))

    def get_active_window_title(self) -> str:
        """获取当前活动窗口标题"""
        if not WINDOW_DETECTION_AVAILABLE:
            return ""

        try:
            hwnd = win32gui.GetForegroundWindow()
            return win32gui.GetWindowText(hwnd)
        except Exception:
            return ""

    def is_target_window_active(self) -> bool:
        """检查目标窗口是否处于激活状态"""
        if not self.enabled:
            return True  # 未启用时，始终返回 True（不限制）

        current_title = self.get_active_window_title()

        # 检查当前窗口标题是否包含目标窗口名称
        return any(target in current_title for target in self.target_window_names)


class MacroEngine:
    """宏引擎 - 负责解析和执行宏指令"""

    # 键盘扫描码映射表
    SCAN_CODE_MAP = {
        30: 'a', 31: 's', 32: 'd', 33: 'f',
        17: 'w', 44: 'z', 45: 'x', 46: 'c',
    }

    def __init__(self, window_monitor: Optional[WindowMonitor] = None):
        self.running = False
        self.stop_flag = False
        self.window_monitor = window_monitor
        self.pressed_keys: List[str] = []  # 追踪按下的键

        # 动作处理器映射
        self._action_handlers: Dict[str, Callable] = {
            'press': self._handle_press,
            'hold': self._handle_hold,

            'delay': self._handle_delay,
            'click': self._handle_click,
            'doubleclick': self._handle_doubleclick,
            'keydown': self._handle_keydown,
            'keyup': self._handle_keyup,
        }

    def parse_delay(self, delay_str: str) -> float:
        """解析延迟时间，支持多种格式：100ms, 0.1s, 1秒"""
        delay_str = str(delay_str).strip().lower()

        try:
            if 'ms' in delay_str:
                # 毫秒：100ms -> 0.1秒
                return float(delay_str.replace('ms', '')) / 1000
            elif 's' in delay_str or '秒' in delay_str:
                # 秒：0.1s, 1秒 -> 秒数
                return float(re.sub(r'[s秒]', '', delay_str))
            else:
                # 纯数字默认为毫秒：100 -> 0.1秒
                return float(delay_str) / 1000
        except ValueError:
            return Config.DEFAULT_DELAY_FALLBACK

    def parse_key(self, key_str: str) -> str:
        """解析按键字符串（已经是英文标准格式）"""
        return str(key_str).strip().lower()

    def parse_key_code(self, key_code: int) -> Optional[str]:
        """解析键盘扫描码（用于 XML 格式）"""
        return self.SCAN_CODE_MAP.get(int(key_code))

    def _handle_press(self, action: Dict[str, Any], repeat_count: int = 1) -> None:
        """处理按键动作"""
        key = action['key']
        for _ in range(repeat_count):
            if self.stop_flag:
                break
            try:
                keyboard.press_and_release(key)
                time.sleep(Config.KEY_PRESS_INTERVAL)
            except Exception:
                pass

    def _handle_hold(self, action: Dict[str, Any], _: int = 1) -> None:
        """处理按住动作"""
        key = action['key']
        duration = action['duration']
        try:
            keyboard.press(key)
            time.sleep(duration)
            keyboard.release(key)
        except Exception:
            pass

    def _handle_delay(self, action: Dict[str, Any], _: int = 1) -> None:
        """处理延迟动作（可中断）"""
        duration = action['duration']
        # 将长延迟拆分成多个短延迟，以便可以中断
        elapsed = 0.0

        while elapsed < duration and not self.stop_flag:
            remaining = duration - elapsed
            sleep_time = min(Config.DELAY_CHECK_INTERVAL, remaining)
            time.sleep(sleep_time)
            elapsed += sleep_time

    def _handle_click(self, action: Dict[str, Any], _: int = 1) -> None:
        """处理鼠标点击动作"""
        button = action.get('button', 'left')
        mouse.click(button)

    def _handle_doubleclick(self, action: Dict[str, Any], _: int = 1) -> None:
        """处理鼠标双击动作"""
        button = action.get('button', 'left')
        mouse.double_click(button)

    def _handle_keydown(self, action: Dict[str, Any], _: int = 1) -> None:
        """处理按下键（不释放）"""
        key = action['key']
        try:
            keyboard.press(key)
            # 追踪按下的键
            if key not in self.pressed_keys:
                self.pressed_keys.append(key)
        except Exception:
            pass

    def _handle_keyup(self, action: Dict[str, Any], _: int = 1) -> None:
        """处理释放键"""
        key = action['key']
        try:
            keyboard.release(key)
            # 从追踪列表中移除
            if key in self.pressed_keys:
                self.pressed_keys.remove(key)
        except Exception:
            pass

    def reset_keys(self, force_release_modifiers: bool = False) -> None:
        """重置所有按住的键

        Args:
            force_release_modifiers: 是否强制释放所有修饰键（用于紧急情况）
        """
        # 释放追踪的按键
        for key in self.pressed_keys[:]:  # 使用副本遍历
            try:
                keyboard.release(key)
            except Exception:
                pass
        self.pressed_keys.clear()

        # 强制释放修饰键（防止输入法干扰导致按键卡住）
        if force_release_modifiers:
            for modifier_key in Config.FORCE_RELEASE_MODIFIER_KEYS:
                try:
                    keyboard.release(modifier_key)
                except Exception:
                    pass

    def execute_action(self, action: Dict[str, Any], repeat_count: int = 1) -> None:
        """执行单个动作"""
        if self.stop_flag:
            return

        action_type = action['type']
        handler = self._action_handlers.get(action_type)

        if handler:
            handler(action, repeat_count)

    def execute_macro(self, macro: Dict[str, Any], repeat_mode: str = 'once',
                     key_state_checker: Optional[Callable[[], bool]] = None,
                     additional_keys: Optional[List[str]] = None) -> None:
        """执行宏序列

        Args:
            macro: 宏配置
            repeat_mode: 重复模式 ('once', 'loop', 'hold')
            key_state_checker: 按键状态检查函数，返回 True 表示继续执行，False 表示停止
            additional_keys: 附加按键列表，这些键在 reset_keys 后需要重新按下
        """
        self.running = True
        self.stop_flag = False
        actions = macro['actions']
        start_actions = macro.get('start_actions', [])
        reset_keys_enabled = macro.get('reset_keys', False)
        skip_window_check = macro.get('skip_window_check', False)

        # 如果启用了"重置按键"，在宏开始执行前立即释放所有键
        if reset_keys_enabled:
            self.reset_keys(force_release_modifiers=True)

            # 重新按下附加按键（避免附加按键被 reset_keys 释放）
            if additional_keys:
                for key in additional_keys:
                    try:
                        keyboard.press(key)
                    except Exception:
                        pass

        # 执行起始动作（只在第一次触发时执行）
        if start_actions:
            for action in start_actions:
                if self.stop_flag:
                    break
                self.execute_action(action)

        def run_actions():
            """执行一次所有动作"""
            for action in actions:
                if self.stop_flag:
                    return False

                # 检查按键状态（用于"按住时"模式）
                if key_state_checker and not key_state_checker():
                    return False

                # 窗口焦点检查（如果启用）
                if not skip_window_check:
                    if self.window_monitor and not self.window_monitor.is_target_window_active():
                        log.warning(_msg('warn_window_lost'))
                        self.reset_keys(force_release_modifiers=True)  # 窗口切换时强制释放所有按键
                        return False

                self.execute_action(action)
            return True

        # 根据模式执行
        if repeat_mode == 'once':
            run_actions()
        else:  # loop 或 hold
            while not self.stop_flag:
                # 周期性检查按键状态
                if key_state_checker and not key_state_checker():
                    break

                if not run_actions():
                    break

        # 执行结束动作（如果配置了）
        finish_actions = macro.get('finish_actions', [])
        if finish_actions and not self.stop_flag:
            for action in finish_actions:
                if self.stop_flag:
                    break
                self.execute_action(action)

        self.running = False

    def stop_macro(self) -> None:
        """停止当前宏并释放所有按住的键"""
        self.stop_flag = True
        self.reset_keys()


class MacroParser:
    """配置文件解析器 - 支持文本格式和 XML 格式"""

    def __init__(self, window_monitor: Optional[WindowMonitor] = None):
        self.engine = MacroEngine(window_monitor)

        # 配置项处理器映射（依赖实例方法）
        self._config_handlers = {
            'trigger_key': self._handle_trigger_key,
            'loop': self._handle_loop,
            'repeat': self._handle_repeat,
            'reset_keys': self._handle_reset_keys,
            'additional_keys': self._handle_additional_keys,
            'default_delay': self._handle_default_delay,
            'skip_window_check': self._handle_skip_window_check,
        }

        # 动作解析器映射（合并实例方法和静态映射）
        self._action_parsers = {
            # 键盘动作（依赖实例方法）
            'press': self._parse_press_action,
            'hold': self._parse_hold_action,
            'keydown': self._parse_keydown_action,
            'keyup': self._parse_keyup_action,
            'delay': self._parse_delay_action,
        }
        # 添加鼠标动作（从 LanguageMapping 获取）
        self._action_parsers.update(LanguageMapping.MOUSE_ACTIONS)


    def _translate(self, text: str) -> str:
        """翻译中文和英文同义词为标准英文"""
        # 先尝试中文映射
        result = LanguageMapping.ZH_CN.get(text, text)
        # 再尝试英文同义词映射
        result = LanguageMapping.EN_SYNONYMS.get(result, result)
        return result

    def _handle_trigger_key(self, value: str, macro: Dict[str, Any]) -> None:
        """处理触发键配置"""
        macro['trigger_key'] = value.lower()

    def _handle_loop(self, value: str, macro: Dict[str, Any]) -> None:
        """处理循环配置"""
        translated_value = self._translate(value).lower()
        if translated_value == 'true':
            macro['repeat_mode'] = 'loop'
        elif translated_value == 'false':
            macro['repeat_mode'] = 'once'

    def _handle_repeat(self, value: str, macro: Dict[str, Any]) -> None:
        """处理重复配置"""
        if self._translate(value) == 'hold':
            macro['repeat_mode'] = 'hold'
        elif value.isdigit():
            macro['repeat_count'] = int(value)

    def _handle_reset_keys(self, value: str, macro: Dict[str, Any]) -> None:
        """处理重置按键配置"""
        translated_value = self._translate(value).lower()
        macro['reset_keys'] = (translated_value == 'true')

    def _handle_additional_keys(self, value: str, macro: Dict[str, Any]) -> None:
        """处理附加按键配置"""
        # 分割键（支持逗号、中文顿号或空格分隔）
        keys_str = value.replace(',', ' ').replace('、', ' ')
        keys = [self.engine.parse_key(self._translate(k.strip()))
                for k in keys_str.split() if k.strip()]
        macro['additional_keys'] = keys if keys else []

    def _handle_default_delay(self, value: str, macro: Dict[str, Any]) -> None:
        """处理默认延迟配置"""
        duration = self.engine.parse_delay(value)
        macro['default_delay'] = duration

    def _handle_skip_window_check(self, value: str, macro: Dict[str, Any]) -> None:
        """处理跳过窗口检测配置"""
        translated_value = self._translate(value).lower()
        macro['skip_window_check'] = (translated_value == 'true')

    def _parse_press_action(self, parts: List[str]) -> Optional[List[Dict[str, Any]]]:
        """解析按键动作"""
        if len(parts) >= 2:
            # 合并所有参数（可能包含逗号分隔的键）
            keys_str = ' '.join(parts[1:])
            # 移除逗号和中文顿号，按空格分割
            keys_str = keys_str.replace(',', ' ').replace('、', ' ')
            keys = [k.strip() for k in keys_str.split() if k.strip()]

            # 如果有多个键，返回多个动作
            if len(keys) > 1:
                return [{'type': 'press', 'key': self.engine.parse_key(self._translate(k))}
                       for k in keys]
            elif len(keys) == 1:
                return [{'type': 'press', 'key': self.engine.parse_key(self._translate(keys[0]))}]
        return None

    def _parse_hold_action(self, parts: List[str]) -> Optional[Dict[str, Any]]:
        """解析按住动作"""
        if len(parts) >= 3:
            key = self.engine.parse_key(self._translate(parts[1]))
            duration = self.engine.parse_delay(parts[2])
            return {'type': 'hold', 'key': key, 'duration': duration}
        return None

    def _parse_delay_action(self, parts: List[str]) -> Optional[Dict[str, Any]]:
        """解析延迟动作"""
        if len(parts) >= 2:
            duration = self.engine.parse_delay(parts[1])
            return {'type': 'delay', 'duration': duration}
        return None

    def _parse_keydown_action(self, parts: List[str]) -> Optional[Dict[str, Any]]:
        """解析按住动作（按下不松开）"""
        if len(parts) >= 2:
            key = self.engine.parse_key(self._translate(parts[1]))
            return {'type': 'keydown', 'key': key}
        return None

    def _parse_keyup_action(self, parts: List[str]) -> Optional[Dict[str, Any]]:
        """解析松开动作"""
        if len(parts) >= 2:
            key = self.engine.parse_key(self._translate(parts[1]))
            return {'type': 'keyup', 'key': key}
        return None


    def parse_text_format(self, content: str) -> List[Dict[str, Any]]:
        """解析简洁的文本格式"""
        macros = []
        current_macro = None
        current_action_section = 'actions'  # 当前正在解析的动作区域

        for line in content.split('\n'):
            line = line.strip()

            # 跳过空行和注释
            if not line or line.startswith('#'):
                continue

            # 新的宏定义
            if line.startswith('[') and line.endswith(']'):
                if current_macro:
                    macros.append(current_macro)

                current_macro = {
                    'name': line[1:-1].strip(),
                    'trigger_key': None,
                    'repeat_mode': 'once',
                    'actions': [],
                    'start_actions': [],  # 起始动作列表
                    'finish_actions': []  # 结束动作列表
                }
                current_action_section = 'actions'  # 重置为默认动作
                continue

            if not current_macro:
                continue

            # 解析配置项
            if '=' in line:
                # 检查是否是动作区域配置项
                key_part = line.split('=', 1)[0].strip()
                translated_key = self._translate(key_part)

                if translated_key == 'start_actions':
                    # 切换到解析起始动作模式
                    current_action_section = 'start_actions'
                elif translated_key == 'finish_actions':
                    # 切换到解析结束动作模式
                    current_action_section = 'finish_actions'
                elif translated_key == 'actions':
                    # 切换回解析普通动作模式
                    current_action_section = 'actions'
                else:
                    # 其他配置项
                    self._parse_config_line(line, current_macro)
                    current_action_section = 'actions'  # 遇到其他配置项时，重置为普通动作模式
            else:
                # 解析动作指令（支持一行多个指令，用多个空格分隔）
                actions = self._parse_action_line(line)
                if actions:
                    # 根据当前区域添加到对应的动作列表
                    current_macro[current_action_section].extend(actions)

        if current_macro:
            # 在解析完所有动作后，插入默认延迟
            self._insert_default_delays(current_macro)
            macros.append(current_macro)

        return macros

    def _insert_default_delays(self, macro: Dict[str, Any]) -> None:
        """在每个非延迟动作后插入默认延迟"""
        default_delay = macro.get('default_delay')
        if not default_delay:
            return

        # 处理所有动作区域（起始、普通、结束）
        for action_key in ['start_actions', 'actions', 'finish_actions']:
            actions = macro.get(action_key, [])
            if not actions:
                continue

            new_actions = []
            for i, action in enumerate(actions):
                new_actions.append(action)

                # 如果当前动作不是delay，且不是最后一个动作，插入默认延迟
                if action['type'] != 'delay' and i < len(actions) - 1:
                    # 检查下一个动作是否已经是delay
                    next_action = actions[i + 1]
                    if next_action['type'] != 'delay':
                        # 插入默认延迟
                        new_actions.append({'type': 'delay', 'duration': default_delay})

            macro[action_key] = new_actions

    def _parse_config_line(self, line: str, macro: Dict[str, Any]) -> None:
        """解析配置行"""
        try:
            if '=' not in line:
                log.warning(_msg('warn_config_format', line))
                return

            key, value = line.split('=', 1)
            key = self._translate(key.strip())
            value = value.strip()

            # 使用处理器映射
            handler = self._config_handlers.get(key)
            if handler:
                handler(value, macro)
            else:
                # 未知的配置项，给出警告但不中断
                if key and not key.startswith('_'):  # 跳过空键和内部键
                    log.warning(_msg('warn_unknown_config', key, macro['name']))
        except Exception as e:
            log.error(_msg('warn_parse_failed', line))
            log.error(_msg('warn_error', e))

    def _parse_action_line(self, line: str) -> List[Dict[str, Any]]:
        """解析单行动作指令（支持一行多个指令）"""
        # 移除注释
        line = line.split('#')[0].strip()
        if not line:
            return []

        parts = line.split()
        if not parts:
            return []

        actions = []
        i = 0

        while i < len(parts):
            # 翻译并检查是否是动作关键字
            action_type = self._translate(parts[i])
            parser = self._action_parsers.get(action_type)

            if parser:
                # 找到下一个动作关键字的位置
                next_action_idx = None
                for j in range(i + 1, len(parts)):
                    if self._translate(parts[j]) in self._action_parsers:
                        next_action_idx = j
                        break

                # 提取当前动作的所有参数
                if next_action_idx:
                    action_parts = parts[i:next_action_idx]
                else:
                    action_parts = parts[i:]

                # 解析动作
                result = parser(action_parts)
                if result:
                    # 处理返回列表或单个动作的情况
                    if isinstance(result, list):
                        actions.extend(result)
                    else:
                        actions.append(result)

                # 移动到下一个动作
                i = next_action_idx if next_action_idx else len(parts)
            else:
                i += 1

        return actions

    def parse_xml_format(self, content: str) -> List[Dict[str, Any]]:
        """解析 XML 格式 """
        try:
            root = ET.fromstring(content)
            macros = []

            for macro_elem in root.findall('.//DefaultMacro'):
                repeat_type = macro_elem.findtext('.//RepeatType', '0')
                keydown_syntax = macro_elem.findtext('.//KeyDown/Syntax', '')

                macros.append({
                    'name': macro_elem.findtext('Major', _msg('unnamed_macro')),
                    'description': macro_elem.findtext('Description', ''),
                    'trigger_key': None,
                    'repeat_mode': LanguageMapping.XML_REPEAT_MODE.get(repeat_type, 'once'),
                    'actions': self._parse_xml_syntax(keydown_syntax)
                })

            return macros
        except Exception as e:
            log.error(_msg('parse_xml_failed', e))
            return []

    def _parse_xml_syntax(self, syntax: str) -> List[Dict[str, Any]]:
        """解析 XML 格式的按键语法（简单直接）"""
        actions = []

        for line in syntax.strip().split('\n'):
            parts = line.strip().split()
            if not parts or len(parts) < 2:
                continue

            cmd, *args = parts

            if cmd in ('KeyDown', 'KeyUp'):
                key = self.engine.parse_key_code(int(args[0]))
                if key:
                    action_type = 'keydown' if cmd == 'KeyDown' else 'keyup'
                    actions.append({'type': action_type, 'key': key})

            elif cmd == 'Delay':
                actions.append({'type': 'delay', 'duration': int(args[0]) / 1000})

            elif cmd == 'LeftDown':
                actions.append({'type': 'click', 'button': 'left'})

            elif cmd == 'RightDown':
                actions.append({'type': 'click', 'button': 'right'})

        return actions

    def load_file(self, filepath: str) -> List[Dict[str, Any]]:
        """加载配置文件"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()

            # 判断文件格式
            return (self.parse_xml_format(content) if content.strip().startswith('<')
                    else self.parse_text_format(content))
        except Exception as e:
            log.error(_msg('load_file_failed', e))
            return []


class MacroRunner:
    """宏运行器 - 负责热键监听和宏执行"""

    def __init__(self, config_file: str, target_window_names: Optional[List[str]] = None):
        """
        初始化宏运行器

        Args:
            config_file: 配置文件路径
            target_window_names: 目标窗口名称列表（如 ['Diablo IV']），None 表示不限制窗口
        """
        self.window_monitor = WindowMonitor(target_window_names)
        self.parser = MacroParser(self.window_monitor)
        self.engine = self.parser.engine
        self.config_file = config_file
        self.macros: List[Dict[str, Any]] = []
        self.hotkey_map: Dict[str, Dict[str, Any]] = {}
        self.current_thread: Optional[threading.Thread] = None
        self.paused = False

        # Esc 双击检测
        self.last_esc_time = 0
        self.esc_double_click_interval = Config.ESC_DOUBLE_CLICK_INTERVAL

        # 附加按键追踪
        self.macro_additional_keys: Dict[str, List[str]] = {}
        self.additional_keys_lock = threading.Lock()

        # 按键输入缓冲队列
        self.input_buffer: deque = deque(maxlen=Config.INPUT_BUFFER_SIZE)
        self.buffer_lock = threading.Lock()

        # 控制热键防抖 - 记录上次按下时间
        self.last_pause_time = 0
        self.last_reload_time = 0
        self.last_exit_time = 0
        self.control_key_debounce = Config.CONTROL_KEY_DEBOUNCE

        # 追踪每个宏最后一次触发时间（宏名称 -> 时间戳）
        self.last_macro_trigger_time: Dict[str, float] = {}

        # 追踪当前按住的触发键（用于"按住时"模式的状态检查）
        self.held_trigger_keys: set = set()
        self.held_keys_lock = threading.Lock()
        self.current_hold_macro_name: Optional[str] = None

        self.load_config()

    def load_config(self) -> None:
        """加载配置"""
        self.macros = self.parser.load_file(self.config_file)

        if not self.macros:
            log.warning(_msg('no_macros_found'))
            return

        # 只显示总数，不显示每个宏的详细信息
        log.success(f"加载 {len(self.macros)} 个宏", use_icon=False)

        self._build_hotkey_map()

    def _build_hotkey_map(self) -> None:
        """构建热键映射"""
        self.hotkey_map = {}
        for macro in self.macros:
            trigger_key = macro.get('trigger_key')
            if not trigger_key:
                continue

            trigger_key_lower = trigger_key.lower()
            if trigger_key_lower in self.hotkey_map:
                existing_macro = self.hotkey_map[trigger_key_lower]
                log.warning(_msg('warn_duplicate_key', trigger_key))
                log.warning(_msg('warn_key_override', existing_macro['name'], macro['name']))

            self.hotkey_map[trigger_key_lower] = macro

    def setup_hotkeys(self) -> None:
        """设置热键"""
        keyboard.unhook_all()

        # 系统热键 - 使用 on_press_key 确保最高优先级和可靠性
        keyboard.on_press_key(Config.HOTKEY_EXIT, lambda _: self._handle_exit_key())
        keyboard.on_press_key(Config.HOTKEY_RELOAD, lambda _: self._handle_reload_key())
        keyboard.on_press_key(Config.HOTKEY_PAUSE, lambda _: self._handle_pause_key())

        # 紧急停止热键（双击 Esc）
        keyboard.on_press_key('esc', lambda _: self.emergency_stop())

        # 宏热键
        for key_str, macro in self.hotkey_map.items():
            repeat_mode = macro.get('repeat_mode', 'once')
            has_additional_keys = bool(macro.get('additional_keys'))

            # hold 模式或有附加按键的宏，需要监听释放事件
            if repeat_mode == 'hold' or has_additional_keys:
                keyboard.on_press_key(key_str, lambda _, m=macro: self._on_trigger_press(m))
                keyboard.on_release_key(key_str, lambda _, m=macro: self._on_trigger_release(m))
            else:
                # 其他模式且无附加按键，直接启动
                keyboard.on_press_key(key_str, lambda _, m=macro: self.start_macro(m))

    def _handle_pause_key(self) -> None:
        """处理暂停热键"""
        current_time = time.time()
        if current_time - self.last_pause_time > self.control_key_debounce:
            self.last_pause_time = current_time
            self.toggle_pause()

    def _handle_reload_key(self) -> None:
        """处理重新加载热键"""
        current_time = time.time()
        if current_time - self.last_reload_time > self.control_key_debounce:
            self.last_reload_time = current_time
            self.reload_config()

    def _handle_exit_key(self) -> None:
        """处理退出热键"""
        current_time = time.time()
        if current_time - self.last_exit_time > self.control_key_debounce:
            self.last_exit_time = current_time
            self.stop()

    def toggle_pause(self) -> None:
        """切换暂停/恢复状态"""
        self.paused = not self.paused

        if self.paused:
            print(f"{Fore.YELLOW}{_msg('pause_on')}{Style.RESET_ALL}")
            self._stop_current_macro()
            self.clear_input_buffer()  # 暂停时清空缓冲队列
            self._force_release_all_keys()  # 强制释放所有按键
        else:
            print(f"{Fore.GREEN}{_msg('pause_off')}{Style.RESET_ALL}")

    def reload_config(self) -> None:
        """重新加载配置文件"""
        self._stop_current_macro()
        self.clear_input_buffer()  # 重新加载时清空缓冲队列
        self._force_release_all_keys()  # 强制释放所有按键
        self.load_config()
        self.setup_hotkeys()
        print(f"{Fore.GREEN}{_msg('config_reloaded')}{Style.RESET_ALL}")

    def start_macro(self, macro: Dict[str, Any]) -> None:
        """启动宏（智能缓冲）"""
        if self.paused:
            return

        macro_name = macro['name']
        current_time = time.time()

        if Config.INPUT_BUFFER_ENABLED and self.engine.running:
            should_buffer = self._should_buffer_macro(macro, current_time)

            if should_buffer:
                with self.buffer_lock:
                    if len(self.input_buffer) >= Config.INPUT_BUFFER_SIZE:
                        return
                    self.input_buffer.append(macro)
                return

        self._stop_current_macro()

        self.last_macro_trigger_time[macro_name] = current_time

        self.current_thread = threading.Thread(
            target=self._execute_macro_with_queue,
            args=(macro,),
            daemon=True
        )
        self.current_thread.start()

    def _should_buffer_macro(self, macro: Dict[str, Any], current_time: float) -> bool:
        """判断是否应该缓冲宏"""
        macro_name = macro['name']
        repeat_mode = macro.get('repeat_mode', 'once')

        # 规则1：如果新宏是"按住时"模式，不缓冲（立即中断）
        if repeat_mode == 'hold':
            return False

        # 规则2：如果新宏在短时间内重复触发，才缓冲（避免误触）
        last_trigger = self.last_macro_trigger_time.get(macro_name, 0)
        time_since_last = current_time - last_trigger

        if time_since_last < Config.RAPID_PRESS_THRESHOLD:
            # 短时间内连按，缓冲
            return True
        else:
            # 不是快速连按，直接中断当前宏
            return False

    def _execute_macro_with_queue(self, macro: Dict[str, Any]) -> None:
        """执行宏，并在完成后处理队列中的下一个宏"""
        macro_name = macro['name']
        trigger_key = macro.get('trigger_key', '')
        repeat_mode = macro.get('repeat_mode', 'once')
        additional_keys = macro.get('additional_keys', [])

        def check_key_still_held() -> bool:
            if repeat_mode != 'hold':
                return True
            with self.held_keys_lock:
                is_held = trigger_key in self.held_trigger_keys
                if not is_held:
                    return False
                return True

        try:
            self.engine.execute_macro(macro, repeat_mode,
                                     key_state_checker=check_key_still_held,
                                     additional_keys=additional_keys)
        finally:
            if Config.INPUT_BUFFER_ENABLED:
                self._process_next_in_queue()

    def _process_next_in_queue(self) -> None:
        """处理队列中的下一个宏"""
        with self.buffer_lock:
            if len(self.input_buffer) > 0:
                next_macro = self.input_buffer.popleft()
                # 静默执行

                # 递归执行下一个宏（使用新线程）
                self.current_thread = threading.Thread(
                    target=self._execute_macro_with_queue,
                    args=(next_macro,),
                    daemon=True
                )
                self.current_thread.start()

    def clear_input_buffer(self) -> None:
        """清空输入缓冲队列"""
        with self.buffer_lock:
            self.input_buffer.clear()
            # 静默清空

    def _stop_current_macro(self) -> None:
        """停止当前运行的宏"""
        if self.engine.running:
            self.engine.stop_macro()
            if self.current_thread:
                self.current_thread.join(timeout=Config.THREAD_JOIN_TIMEOUT)

    def _force_release_all_keys(self) -> None:
        """强制释放所有按键（包括附加按键和修饰键）"""
        self.engine.reset_keys(force_release_modifiers=True)
        with self.additional_keys_lock:
            for keys in self.macro_additional_keys.values():
                for key in keys:
                    try:
                        keyboard.release(key)
                    except Exception:
                        pass
            self.macro_additional_keys.clear()
        with self.held_keys_lock:
            self.held_trigger_keys.clear()
            self.current_hold_macro_name = None

    def _on_trigger_press(self, macro: Dict[str, Any]) -> None:
        """触发键按下时的处理（按住附加按键并启动宏）"""
        macro_name = macro['name']
        trigger_key = macro.get('trigger_key', '')
        additional_keys = macro.get('additional_keys', [])

        with self.held_keys_lock:
            self.held_trigger_keys.add(trigger_key)
            if macro.get('repeat_mode') == 'hold':
                self.current_hold_macro_name = macro_name

        if additional_keys:
            with self.additional_keys_lock:
                self.macro_additional_keys[macro_name] = additional_keys.copy()
                for key in additional_keys:
                    try:
                        keyboard.press(key)
                    except Exception:
                        pass

        self.start_macro(macro)

    def _on_trigger_release(self, macro: Dict[str, Any]) -> None:
        """触发键松开时的处理（根据模式决定是否停止宏和释放附加按键）"""
        macro_name = macro['name']
        trigger_key = macro.get('trigger_key', '')
        repeat_mode = macro.get('repeat_mode', 'once')

        with self.held_keys_lock:
            self.held_trigger_keys.discard(trigger_key)
            if self.current_hold_macro_name == macro_name:
                self.current_hold_macro_name = None

        self._release_additional_keys(macro_name)

        if repeat_mode == 'hold':
            if self.engine.running:
                self.engine.stop_macro()

    def _release_additional_keys(self, macro_name: str) -> None:
        """释放指定宏的附加按键"""
        with self.additional_keys_lock:
            if macro_name in self.macro_additional_keys:
                keys = self.macro_additional_keys[macro_name]
                for key in keys:
                    try:
                        keyboard.release(key)
                    except Exception:
                        pass
                del self.macro_additional_keys[macro_name]

    def emergency_stop(self) -> None:
        """紧急停止所有宏（双击 Esc）"""
        current_time = time.time()

        # 检查是否在双击间隔内
        if current_time - self.last_esc_time <= self.esc_double_click_interval:
            # 确认双击，执行紧急停止
            print(f"{Fore.RED}{_msg('emergency_stopped')}{Style.RESET_ALL}")
            self._stop_current_macro()
            self.clear_input_buffer()
            self.engine.reset_keys(force_release_modifiers=True)

            with self.additional_keys_lock:
                for keys in self.macro_additional_keys.values():
                    for key in keys:
                        try:
                            keyboard.release(key)
                        except Exception:
                            pass
                self.macro_additional_keys.clear()

            with self.held_keys_lock:
                self.held_trigger_keys.clear()
                self.current_hold_macro_name = None

            self.last_esc_time = 0
        else:
            # 第一次按下，记录时间
            self.last_esc_time = current_time

    def start(self) -> None:
        """启动宏运行器"""
        # 简洁的启动提示
        print(f"\n{Fore.GREEN}{_msg('app_start')}{Style.RESET_ALL}  "
              f"{Style.DIM}{_msg('app_hints', Config.HOTKEY_PAUSE, Config.HOTKEY_RELOAD, Config.HOTKEY_EXIT)}{Style.RESET_ALL}")

        self.setup_hotkeys()
        try:
            while True:
                time.sleep(0.1)
        except SystemExit:
            pass

    def stop(self) -> None:
        """停止宏运行器"""
        print(f"\n{Fore.CYAN}{_msg('app_exiting')}{Style.RESET_ALL}")
        self._stop_current_macro()
        self._force_release_all_keys()  # 强制释放所有按键
        keyboard.unhook_all()
        sys.exit(0)


def main() -> None:
    """主函数"""
    print(f"\n{Style.BRIGHT}{Config.APP_NAME} v{Config.APP_VERSION}{Style.RESET_ALL}")

    # 显示警告
    if not WINDOW_DETECTION_AVAILABLE:
        print(f"{Fore.YELLOW}{_msg('pywin32_not_installed')}{Style.RESET_ALL}")

    if not os.path.exists(Config.CONFIG_FILE):
        print(f"{Fore.RED}{_msg('config_not_found', Config.CONFIG_FILE)}{Style.RESET_ALL}")
        input(_msg('press_enter_exit'))
        return

    runner = MacroRunner(Config.CONFIG_FILE, Config.TARGET_WINDOWS)

    try:
        runner.start()
    except KeyboardInterrupt:
        print(f"\n{Fore.CYAN}{_msg('app_stopped')}{Style.RESET_ALL}")


if __name__ == '__main__':
    main()
