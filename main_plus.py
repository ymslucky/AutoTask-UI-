import sys
import json
import threading
import time
from collections import deque
from copy import deepcopy
from datetime import datetime, date, timedelta

import openpyxl, random, itertools
import pyautogui
import pyperclip

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QPushButton, QGroupBox, QLabel, QLineEdit,
    QComboBox, QTimeEdit, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView, QFormLayout, QSpinBox,
    QListWidgetItem, QCheckBox, QMenu, QFrame, QStyle,
    QSplitter, QSizePolicy, QDialog, QDialogButtonBox,
    QGridLayout, QFileDialog, QMessageBox, QDoubleSpinBox,
    QTextEdit, QPlainTextEdit, QSystemTrayIcon, QScrollArea, QInputDialog, QDateEdit, QDateTimeEdit, QWidgetAction,
    QButtonGroup
)
from PySide6.QtCore import Qt, QTime, QSize, QSettings, Signal, QObject, QTimer, QPointF, QRectF, QDate, QDateTime, \
    QPoint, QRect, QThread, Slot
from PySide6.QtGui import QIcon, QAction, QFont, QPalette, QColor, QLinearGradient, QTextCursor, QKeySequence, QPixmap, \
    QBrush, QPainterPath, QPainter, QPen, QMouseEvent, QIntValidator, QCursor, QKeyEvent, QFontMetrics, QScreen

from pathlib import Path

from pynput import keyboard
from pynput.keyboard import KeyCode, Key

from typing import Any, TypedDict

import os
import logging
from typing import Optional, List, Dict
from openai import OpenAI
# 工具函数
def resource_path(relative_path: str) -> str:
    """打包 / 开发环境下通用的资源路径解析"""
    try:
        base_path = sys._MEIPASS           # PyInstaller 运行时
    except AttributeError:
        base_path = os.path.abspath(".")   # 开发环境
    return os.path.join(base_path, relative_path)
# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class StrongRandomSelector:
    def __init__(self, cells, window_size=30):
        """
        初始化强随机选择器
        :param cells: 原始 cell 列表
        :param window_size: 滑动窗口大小，默认为 30
        """
        self.cells = list(cells)  # 原始数据副本
        self.window_size = window_size
        self.recent_cells = deque(maxlen=window_size)  # 滑动窗口队列

    def select_next(self):
        """
        选择下一个强随机的 cell
        :return: 选中的 cell
        """
        # 从原始数据中筛选出不在滑动窗口中的候选池
        candidates = [cell for cell in self.cells if cell not in self.recent_cells]

        # 边界处理：如果候选池为空，清空队列或允许重复
        if not candidates:
            print("[WARNING] 候选池为空，清空滑动窗口以继续选择")
            self.recent_cells.clear()
            candidates = self.cells  # 允许重复

        # 从候选池中随机选择一个 cell
        selected_cell = random.choice(candidates)

        # 将选中的 cell 加入滑动窗口队列
        self.recent_cells.append(selected_cell)

        return selected_cell



# 在 main_plus.py 中添加以下类
class CoordinatePickerOverlay(QDialog):
    """
    坐标拾取覆盖层，用于获取鼠标位置坐标
    """
    coordinate_selected = Signal(tuple)  # 发送选中的坐标 (x, y)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)  # 关键：不遮挡鼠标
        self.setModal(True)

        # 关键修复：设置窗口可接受焦点
        self.setFocusPolicy(Qt.StrongFocus)
        self.setFocus()

        # 获取屏幕信息和缩放比例
        self.screen = QApplication.primaryScreen()
        self.screen_geometry = self.screen.geometry()
        self.device_pixel_ratio = self.screen.devicePixelRatio()

        # 调试信息
        print(f"[DEBUG] 屏幕尺寸: {self.screen_geometry.width()}x{self.screen_geometry.height()}")
        print(f"[DEBUG] 设备像素比例: {self.device_pixel_ratio}")

        self.setGeometry(self.screen_geometry)

        # 创建坐标显示标签
        self.coord_label = QLabel(self)
        self.coord_label.setStyleSheet("""
            background-color: rgba(0, 0, 0, 200);
            color: white;
            padding: 8px 12px;
            border-radius: 6px;
            font-family: Consolas, monospace;
            font-size: 13px;
            border: 2px solid rgba(255, 255, 255, 150);
            font-weight: bold;
        """)
        self.coord_label.hide()

        # 创建提示标签
        self.tip_label = QLabel("按 Enter 确认坐标，按 Esc 取消 | 左键点击也可确认", self)
        self.tip_label.setStyleSheet("""
            background-color: rgba(0, 0, 0, 200);
            color: #FFA500;
            padding: 10px 16px;
            border-radius: 6px;
            font-family: Microsoft YaHei, sans-serif;
            font-size: 12px;
            border: 2px solid rgba(255, 165, 0, 150);
            font-weight: bold;
        """)
        self.tip_label.hide()

        # 定时器用于更新坐标显示
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_position)
        self.timer.start(16)  # ~60fps

        # 跟踪鼠标位置
        self.current_pos = QPoint(0, 0)
        self.raw_pos = QPoint(0, 0)  # 原始坐标

        # 显示提示信息
        QTimer.singleShot(100, self.show_tip)

    def get_scaled_coordinates(self, pos):
        """
        获取缩放校正后的坐标
        返回调整后的坐标和原始坐标
        """
        raw_x, raw_y = pos.x(), pos.y()

        # 方法1: 使用设备像素比例校正
        scaled_x = int(raw_x * self.device_pixel_ratio)
        scaled_y = int(raw_y * self.device_pixel_ratio)

        # 方法2: 备用方法 - 使用屏幕虚拟大小
        virtual_geometry = self.screen.virtualGeometry()
        if virtual_geometry.width() != self.screen_geometry.width():
            scale_factor = virtual_geometry.width() / self.screen_geometry.width()
            scaled_x = int(raw_x * scale_factor)
            scaled_y = int(raw_y * scale_factor)

        return (scaled_x, scaled_y), (raw_x, raw_y)

    def update_position(self):
        """更新鼠标位置显示"""
        mouse_pos = QCursor.pos()
        self.raw_pos = mouse_pos

        # 获取校正后的坐标
        scaled_coords, raw_coords = self.get_scaled_coordinates(mouse_pos)
        self.current_pos = QPoint(scaled_coords[0], scaled_coords[1])

        # 更新坐标标签文本
        coord_text = f"坐标: {scaled_coords[0]}, {scaled_coords[1]}"
        coord_text += f"\n原始: {raw_coords[0]}, {raw_coords[1]}"
        coord_text += f"\n缩放: {self.device_pixel_ratio:.1f}x"

        self.coord_label.setText(coord_text)
        self.coord_label.adjustSize()

        # 标签定位（避免超出屏幕边界）
        label_x = mouse_pos.x() + 25
        label_y = mouse_pos.y() + 25

        if label_x + self.coord_label.width() > self.screen_geometry.width():
            label_x = mouse_pos.x() - self.coord_label.width() - 15
        if label_y + self.coord_label.height() > self.screen_geometry.height():
            label_y = mouse_pos.y() - self.coord_label.height() - 15

        self.coord_label.move(label_x, label_y)
        self.coord_label.show()


    def show_tip(self):
        """显示操作提示"""
        self.tip_label.show()
        self.tip_label.adjustSize()

        # 将提示标签定位在屏幕中央底部
        tip_x = (self.screen_geometry.width() - self.tip_label.width()) // 2
        tip_y = self.screen_geometry.height() - 100
        self.tip_label.move(tip_x, tip_y)


    def showEvent(self, event):
        """窗口显示时自动获取焦点"""
        super().showEvent(event)
        self.setFocus()
        self.activateWindow()
        # 确保窗口在最前面
        self.raise_()
        self.timer.start(16)

    def hideEvent(self, event):
        """窗口隐藏时停止定时器"""
        super().hideEvent(event)
        self.timer.stop()

    def keyPressEvent(self, event):
        """键盘事件处理"""
        if event.key() in (Qt.Key_Enter, Qt.Key_Return):
            print(f"[DEBUG] 确认坐标 - 原始: ({self.raw_pos.x()}, {self.raw_pos.y()}), "
                  f"校正: ({self.current_pos.x()}, {self.current_pos.y()})")
            self.coordinate_selected.emit((self.current_pos.x(), self.current_pos.y()))
            self.accept()
        elif event.key() == Qt.Key_Escape:
            self.reject()
        elif event.key() == Qt.Key_Space:
            # 空格键切换放大镜显示
            self.magnifier.setVisible(not self.magnifier.isVisible())
        else:
            super().keyPressEvent(event)

    def exec_(self):
        """重写exec_方法确保焦点正确设置"""
        self.setFocus()
        self.activateWindow()
        self.raise_()
        return super().exec_()

    def mousePressEvent(self, event):
        """鼠标点击时也确认坐标"""
        if event.button() == Qt.LeftButton:
            print(f"[DEBUG] 鼠标确认坐标 - 原始: ({self.raw_pos.x()}, {self.raw_pos.y()}), "
                  f"校正: ({self.current_pos.x()}, {self.current_pos.y()})")
            self.coordinate_selected.emit((self.current_pos.x(), self.current_pos.y()))
            self.accept()
        else:
            super().mousePressEvent(event)



class AITestDialog(QDialog):
    """AI 测试对话框，支持 Kimi 和豆包"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🤖 AI 测试")
        self.setModal(True)
        self.resize(650, 550)

        # 初始化配置管理器
        self.config_manager = ConfigManager()

        # 初始化 ChatBot
        self.chat_bot = None
        self.current_provider = "kimi"
        self.init_chat_bot()

        self.setup_ui()

    def init_chat_bot(self):
        """初始化 ChatBot"""
        try:
            config = self.config_manager.load()

            # 根据可用配置选择默认提供商
            kimi_key = config.get("moonshot_api_key")
            doubao_ak = config.get("volcano_access_key")
            doubao_sk = config.get("volcano_secret_key")
            doubao_endpoint = config.get("ark_endpoint_id")

            # 优先使用 Kimi（如果配置了的话）
            if kimi_key:
                self.chat_bot = ChatBot(
                    provider="kimi",
                    token_json_path="./config/token.json"
                )
                self.current_provider = "kimi"
            # 否则使用豆包（如果配置了的话）
            elif all([doubao_ak, doubao_sk, doubao_endpoint]):
                self.chat_bot = ChatBot(
                    provider="doubao",
                    token_json_path="./config/token.json"
                )
                self.current_provider = "doubao"

        except Exception as e:
            QMessageBox.warning(self, "初始化失败", f"ChatBot 初始化失败: {str(e)}")

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 说明文本
        intro_label = QLabel("AI 测试对话")
        intro_label.setStyleSheet("""
            font-size: 18px; 
            font-weight: bold; 
            color: #2c3e50;
            padding: 5px 0;
            border-bottom: 2px solid #3498db;
            margin-bottom: 2px;
        """)
        layout.addWidget(intro_label)

        # AI 提供商选择
        provider_layout = QHBoxLayout()
        provider_layout.addWidget(QLabel("AI 提供商:"))

        self.provider_combo = QComboBox()
        self.provider_combo.addItems(["Kimi", "豆包"])
        self.provider_combo.setCurrentText("Kimi" if self.current_provider == "kimi" else "豆包")
        self.provider_combo.currentTextChanged.connect(self.on_provider_changed)
        provider_layout.addWidget(self.provider_combo)
        provider_layout.addStretch()

        layout.addLayout(provider_layout)

        # 对话历史区域
        history_group = QGroupBox("对话历史")
        history_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #000000;
                border-radius: 10px;
                margin-top: 1ex;
                padding-top: 2px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subline-position: top center;
                padding: 0 2px;
                background-color: #000000;
                color: white;
                border-radius: 5px;
            }
        """)
        history_layout = QVBoxLayout(history_group)
        history_layout.setContentsMargins(15, 25, 15, 15)

        self.history_display = QPlainTextEdit()
        self.history_display.setReadOnly(True)
        self.history_display.setPlaceholderText("对话历史将显示在这里...")
        self.history_display.setStyleSheet("""
            QPlainTextEdit {
                background-color: white;
                border: 1px solid #bdc3c7;
                border-radius: 8px;
                padding: 2px;
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }
        """)
        # 使用策略扩展，让它占据更多空间
        self.history_display.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # 创建一个容器来更好地控制历史显示区域
        history_container = QWidget()
        history_container_layout = QVBoxLayout(history_container)
        history_container_layout.setContentsMargins(0, 0, 0, 0)
        history_container_layout.addWidget(self.history_display)

        history_layout.addWidget(history_container)
        layout.addWidget(history_group)

        # 用户输入区域
        input_group = QGroupBox("用户输入")
        input_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #000000;
                border-radius: 10px;
                margin-top: 1ex;
                padding-top: 2px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subline-position: top center;
                padding: 0 2px;
                background-color: #000000;
                color: white;
                border-radius: 5px;
            }
        """)
        input_layout = QVBoxLayout(input_group)
        input_layout.setContentsMargins(15, 25, 15, 15)

        self.user_input_edit = QTextEdit()
        self.user_input_edit.setMaximumHeight(180)
        self.user_input_edit.setPlaceholderText("请输入要发送给 AI 的消息...")
        self.user_input_edit.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #bdc3c7;
                border-radius: 8px;
                padding: 10px;
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }
        """)
        input_layout.addWidget(self.user_input_edit)

        # 按钮布局
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)
        button_layout.setContentsMargins(0, 10, 0, 0)

        self.clear_history_btn = QPushButton("🗑️ 清空历史")
        self.clear_history_btn.clicked.connect(self.clear_history)
        self.clear_history_btn.setMinimumWidth(120)
        self.clear_history_btn.setStyleSheet("""
            QPushButton {
                background-color: #95a5a6;
                color: white;
                border: none;
                padding: 10px 16px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7f8c8d;
            }
            QPushButton:pressed {
                background-color: #6c7a7b;
            }
        """)
        button_layout.addWidget(self.clear_history_btn)

        button_layout.addStretch()

        self.send_btn = QPushButton("🚀 发送消息")
        self.send_btn.clicked.connect(self.send_message)
        self.send_btn.setDefault(True)
        self.send_btn.setMinimumWidth(120)
        self.send_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 10px 16px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)
        button_layout.addWidget(self.send_btn)

        input_layout.addLayout(button_layout)
        layout.addWidget(input_group)

        # 状态栏
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("""
            color: #7f8c8d; 
            font-size: 12px;
            padding: 8px 0;
            border-top: 1px solid #ecf0f1;
            font-weight: bold;
        """)
        layout.addWidget(self.status_label)

    def on_provider_changed(self, text):
        """处理 AI 提供商更改"""
        provider = "kimi" if text == "Kimi" else "doubao"
        if provider != self.current_provider:
            self.current_provider = provider
            try:
                self.chat_bot = ChatBot(
                    provider=provider,
                    token_json_path="./config/token.json"
                )
                self.status_label.setText(f"已切换到 {text} 提供商")
                self.clear_history()
            except Exception as e:
                QMessageBox.warning(self, "切换失败", f"切换 AI 提供商失败: {str(e)}")
                # 恢复到之前的提供商
                self.provider_combo.setCurrentText("Kimi" if self.current_provider == "kimi" else "豆包")

    def send_message(self):
        """发送消息到 AI"""
        if not self.chat_bot:
            QMessageBox.warning(self, "错误", "ChatBot 未初始化")
            return

        user_message = self.user_input_edit.toPlainText().strip()
        if not user_message:
            QMessageBox.warning(self, "警告", "请输入消息内容")
            return

        system_prompt = "你是我的朋友，微信语音里很随和。用一句口语化的话回应我"

        # 更新状态
        self.status_label.setText("正在获取 AI 回复...")
        self.send_btn.setEnabled(False)
        QApplication.processEvents()

        try:
            # 发送消息并获取回复
            reply = self.chat_bot.reply(
                message=user_message,
                system=system_prompt if system_prompt else None,
                use_history=True,
                stream=False
            )

            # 更新对话历史
            self.update_history(f"👤 用户: {user_message}")
            self.update_history(f"🤖 AI: {reply}")

            # 清空输入框
            self.user_input_edit.clear()

            self.status_label.setText("回复成功")
        except Exception as e:
            error_msg = f"❌ 错误: {str(e)}"
            self.update_history(error_msg)
            self.status_label.setText("发送失败")
            QMessageBox.critical(self, "发送失败", f"发送消息时出错: {str(e)}")
        finally:
            self.send_btn.setEnabled(True)

    def update_history(self, message):
        """更新对话历史显示"""
        current_text = self.history_display.toPlainText()
        if current_text:
            current_text += "\n" + message
        else:
            current_text = message

        self.history_display.setPlainText(current_text)
        # 滚动到底部
        self.history_display.verticalScrollBar().setValue(
            self.history_display.verticalScrollBar().maximum()
        )

    def clear_history(self):
        """清空对话历史"""
        self.history_display.clear()
        if self.chat_bot:
            self.chat_bot.clear_history()
        self.status_label.setText("历史已清空")


# 在 main_plus.py 文件中添加以下代码

class AITokenConfigDialog(QDialog):
    """AI Token 配置对话框"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("AI Token 配置")
        self.setModal(True)
        self.resize(500, 400)

        # 初始化配置管理器
        self.config_manager = ConfigManager()

        self.setup_ui()
        self.load_config()

    # 替换 AITokenConfigDialog 类中的 setup_ui 方法

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)

        # 说明文本 - 改为可点击链接
        intro_text = QLabel("""
            <p>配置AI服务所需的访问密钥：</p>
            <ul>
                <li><b>Kimi API Key</b>: 用于访问月之暗面的Kimi AI服务 (<a href="https://platform.moonshot.cn/console/api-keys">获取API Key</a>)</li>
                <li><b>豆包 AccessKey/SecretKey</b>: 用于访问字节跳动的豆包AI服务 (<a href="https://www.volcengine.com/product/ark">获取豆包API</a>)</li>
                <li><b>豆包 Endpoint ID</b>: 豆包模型的端点标识</li>
            </ul>
            <p>配置将保存在 <code>./config/token.json</code> 文件中</p>
        """)
        intro_text.setMaximumHeight(120)
        intro_text.setStyleSheet("""
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 6px;
            padding: 10px;
            font-size: 11px;
        """)
        intro_text.setOpenExternalLinks(True)  # 允许打开外部链接
        intro_text.setTextFormat(Qt.RichText)  # 设置为富文本格式
        intro_text.setTextInteractionFlags(Qt.TextBrowserInteraction)  # 允许文本交互
        layout.addWidget(intro_text)

        # Kimi 配置组
        kimi_group = QGroupBox("Kimi 配置")
        kimi_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dcdcdc;
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 15px;
            }
            QGroupBox::title {
                subline-position: top center;
                padding: 0 10px;
            }
        """)
        kimi_layout = QFormLayout(kimi_group)
        kimi_layout.setLabelAlignment(Qt.AlignRight)
        kimi_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        kimi_layout.setHorizontalSpacing(20)
        kimi_layout.setVerticalSpacing(10)

        self.kimi_api_key_edit = QLineEdit()
        self.kimi_api_key_edit.setEchoMode(QLineEdit.Password)
        self.kimi_api_key_edit.setPlaceholderText("请输入 Kimi API Key")
        self.kimi_api_key_edit.setMinimumWidth(200)
        kimi_layout.addRow("API Key:", self.kimi_api_key_edit)

        layout.addWidget(kimi_group)

        # 豆包配置组
        doubao_group = QGroupBox("豆包配置")
        doubao_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dcdcdc;
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 15px;
            }
            QGroupBox::title {
                subline-position: top center;
                padding: 0 10px;
            }
        """)
        doubao_layout = QFormLayout(doubao_group)
        doubao_layout.setLabelAlignment(Qt.AlignRight)
        doubao_layout.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        doubao_layout.setHorizontalSpacing(20)
        doubao_layout.setVerticalSpacing(10)

        self.doubao_ak_edit = QLineEdit()
        self.doubao_ak_edit.setEchoMode(QLineEdit.Password)
        self.doubao_ak_edit.setPlaceholderText("请输入豆包 Access Key")
        self.doubao_ak_edit.setMinimumWidth(200)
        doubao_layout.addRow("Access Key:", self.doubao_ak_edit)

        self.doubao_sk_edit = QLineEdit()
        self.doubao_sk_edit.setEchoMode(QLineEdit.Password)
        self.doubao_sk_edit.setPlaceholderText("请输入豆包 Secret Key")
        self.doubao_sk_edit.setMinimumWidth(200)
        doubao_layout.addRow("Secret Key:", self.doubao_sk_edit)

        self.doubao_endpoint_edit = QLineEdit()
        self.doubao_endpoint_edit.setPlaceholderText("请输入豆包 Endpoint ID")
        self.doubao_endpoint_edit.setMinimumWidth(200)
        doubao_layout.addRow("Endpoint ID:", self.doubao_endpoint_edit)

        layout.addWidget(doubao_group)

        # 按钮布局
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        button_layout.setContentsMargins(0, 10, 0, 0)

        self.load_btn = QPushButton("🔄 重新加载")
        self.load_btn.setStyleSheet("""
            QPushButton {
                background-color: #6c757d;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5a6268;
            }
            QPushButton:pressed {
                background-color: #545b62;
            }
        """)
        self.load_btn.clicked.connect(self.load_config)
        button_layout.addWidget(self.load_btn)

        button_layout.addStretch()

        self.save_btn = QPushButton("💾 保存配置")
        self.save_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:pressed {
                background-color: #1e7e34;
            }
        """)
        self.save_btn.clicked.connect(self.save_config)
        button_layout.addWidget(self.save_btn)

        self.close_btn = QPushButton("❌ 关闭")
        self.close_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
            QPushButton:pressed {
                background-color: #bd2130;
            }
        """)
        self.close_btn.clicked.connect(self.accept)
        button_layout.addWidget(self.close_btn)

        layout.addLayout(button_layout)

    def load_config(self):
        """从配置文件加载配置"""
        try:
            # 加载现有配置
            config = self.config_manager.load()

            # 填充到界面
            self.kimi_api_key_edit.setText(config.get("moonshot_api_key", ""))
            self.doubao_ak_edit.setText(config.get("volcano_access_key", ""))
            self.doubao_sk_edit.setText(config.get("volcano_secret_key", ""))
            self.doubao_endpoint_edit.setText(config.get("ark_endpoint_id", ""))

        except Exception as e:
            QMessageBox.warning(self, "加载失败", f"加载配置时出错: {str(e)}")

    def save_config(self):
        """保存配置到文件"""
        try:
            # 获取界面中的值
            kimi_key = self.kimi_api_key_edit.text().strip()
            doubao_ak = self.doubao_ak_edit.text().strip()
            doubao_sk = self.doubao_sk_edit.text().strip()
            doubao_endpoint = self.doubao_endpoint_edit.text().strip()

            # 检查是否有任何值需要保存
            if not any([kimi_key, doubao_ak, doubao_sk, doubao_endpoint]):
                QMessageBox.information(self, "提示", "没有配置需要保存")
                return

            # 保存配置
            self.config_manager.save(
                moonshot_api_key=kimi_key if kimi_key else None,
                volcano_access_key=doubao_ak if doubao_ak else None,
                volcano_secret_key=doubao_sk if doubao_sk else None,
                ark_endpoint_id=doubao_endpoint if doubao_endpoint else None
            )

            QMessageBox.information(self, "成功", "配置已保存")

        except Exception as e:
            QMessageBox.critical(self, "保存失败", f"保存配置时出错: {str(e)}")


# 配置类型定义
class ConfigDict(TypedDict, total=False):
    moonshot_api_key: str
    volcano_access_key: str
    volcano_secret_key: str
    ark_endpoint_id: str


class ConfigManager:
    """
    配置管理器，用于读写配置文件
    """
    _DEFAULT_DIR = os.path.abspath("./config")
    _DEFAULT_PATH = os.path.join(_DEFAULT_DIR, "token.json")

    def __init__(self, path: Optional[str] = None):
        """
        初始化配置管理器
        :param path: 配置文件路径，默认为 ./config/token.json
        """
        self._path = path or self._DEFAULT_PATH
        self._dir = os.path.dirname(self._path)
        os.makedirs(self._dir, exist_ok=True)

    def save(
            self,
            *,
            moonshot_api_key: Optional[str] = None,
            volcano_access_key: Optional[str] = None,
            volcano_secret_key: Optional[str] = None,
            ark_endpoint_id: Optional[str] = None,
            ensure_ascii: bool = False,
            indent: int = 2
    ) -> None:
        """
        持久化保存配置（按需更新提供的字段）
        :param moonshot_api_key: Kimi API Key
        :param volcano_access_key: 豆包 AccessKey
        :param volcano_secret_key: 豆包 SecretKey
        :param ark_endpoint_id: 豆包 Endpoint ID
        :param ensure_ascii: JSON 是否转义非 ASCII
        :param indent: JSON 缩进空格数
        :raises RuntimeError: 读写失败时抛出
        :raises ValueError: 未提供任何字段时抛出
        """
        updates = {
            "moonshot_api_key": moonshot_api_key,
            "volcano_access_key": volcano_access_key,
            "volcano_secret_key": volcano_secret_key,
            "ark_endpoint_id": ark_endpoint_id,
        }

        if not any(v is not None for v in updates.values()):
            raise ValueError("至少需要提供一个非 None 的配置项")

        # 读取现有配置
        data = self._load_existing_config()

        # 更新配置
        updated = False
        for key, value in updates.items():
            if value is not None:
                data[key] = value
                updated = True

        if not updated:
            raise ValueError("未检测到需要更新的字段")

        self._write_config(data, ensure_ascii, indent)

    def _load_existing_config(self) -> Dict[str, Any]:
        """加载现有配置"""
        if not os.path.exists(self._path):
            return {}

        try:
            with open(self._path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            raise RuntimeError(f"读取配置文件失败: {e}")

    def _write_config(self, data: Dict[str, Any], ensure_ascii: bool, indent: int) -> None:
        """写入配置到文件"""
        try:
            with open(self._path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=ensure_ascii, indent=indent)
        except OSError as e:
            raise RuntimeError(f"写入配置文件失败: {e}")

    def load(self) -> Dict[str, str]:
        """
        读取全部配置
        :return: 配置字典
        :raises RuntimeError: 读取或解析失败时抛出
        """
        return self._load_existing_config()

    def get(
            self,
            key: str,
            default: Optional[str] = None,
            required: bool = False,
    ) -> Optional[str]:
        """
        获取单个配置项
        :param key: 配置键名
        :param default: 默认值
        :param required: 是否为必填项
        :return: 配置值或默认值
        :raises ValueError: 必填项缺失时抛出
        """
        cfg = self.load()
        value = cfg.get(key, default)

        if required and value is None:
            raise ValueError(f"配置项缺失且为必填: {key}")

        return value

    def get_all_keys(self) -> List[str]:
        """
        获取当前配置文件中所有键名
        :return: 键名列表
        """
        cfg = self.load()
        return list(cfg.keys())

    def remove(self, *keys: str, save: bool = True) -> None:
        """
        从配置中移除指定键
        :param keys: 要移除的键名
        :param save: 是否立即写入文件
        :raises RuntimeError: 写入失败时抛出
        """
        if not keys:
            return

        cfg = self.load()
        removed = False

        for key in keys:
            if key in cfg:
                del cfg[key]
                removed = True

        if removed and save:
            # 创建一个空的更新字典来触发保存
            update_dict = {key: None for key in keys}
            self.save(**update_dict)

class ChatBot:
    """聊天机器人客户端"""

    def __init__(
            self,
            provider: str = "kimi",
            kimi_api_key: Optional[str] = None,
            doubao_ak: Optional[str] = None,
            doubao_sk: Optional[str] = None,
            doubao_endpoint_id: Optional[str] = None,
            model: str = "kimi-k2-0905-preview",
            temperature: float = 0.3,
            token_json_path: Optional[str] = None,
    ):
        """
        初始化客户端
        :param provider: 服务提供商 "kimi" | "doubao"
        :param kimi_api_key: Kimi API Key
        :param doubao_ak: 豆包 AccessKey
        :param doubao_sk: 豆包 SecretKey
        :param doubao_endpoint_id: 豆包 Endpoint ID
        :param model: 模型名称
        :param temperature: 温度参数
        :param token_json_path: 自定义 token.json 路径
        """
        self.provider = provider.lower()
        self.model = model
        self.temperature = temperature
        self._client = None
        self._messages: List[Dict[str, str]] = []

        # 获取配置
        config = self._get_config(
            kimi_api_key, doubao_ak, doubao_sk, doubao_endpoint_id, token_json_path
        )

        # 初始化客户端
        self._initialize_client(config)

    def _get_config(
            self,
            kimi_api_key: Optional[str],
            doubao_ak: Optional[str],
            doubao_sk: Optional[str],
            doubao_endpoint_id: Optional[str],
            token_json_path: Optional[str],
    ) -> Dict[str, str]:
        """获取配置信息"""
        config = {}

        # 1) 显式参数优先
        config["moonshot_key"] = self._get_stripped_value(
            kimi_api_key, "MOONSHOT_API_KEY"
        )
        config["ak"] = self._get_stripped_value(doubao_ak, "VOLC_ACCESSKEY")
        config["sk"] = self._get_stripped_value(doubao_sk, "VOLC_SECRETKEY")
        config["endpoint_id"] = self._get_stripped_value(
            doubao_endpoint_id, "ARK_ENDPOINT_ID"
        )

        # 2) 其次尝试从 token.json
        token_cfg = self._load_token_config(token_json_path)

        if not config["moonshot_key"]:
            config["moonshot_key"] = token_cfg.get("moonshot_api_key", "").strip()
        if not config["ak"]:
            config["ak"] = token_cfg.get("volcano_access_key", "").strip()
        if not config["sk"]:
            config["sk"] = token_cfg.get("volcano_secret_key", "").strip()
        if not config["endpoint_id"]:
            config["endpoint_id"] = token_cfg.get("ark_endpoint_id", "").strip()

        return config

    def _get_stripped_value(self, explicit_value: Optional[str], env_var: str) -> str:
        """获取处理后的值（显式参数 > 环境变量）"""
        value = explicit_value or os.getenv(env_var, "")
        return value.strip()

    def _load_token_config(self, token_json_path: Optional[str]) -> Dict[str, str]:
        """从 token.json 加载配置"""
        token_path = token_json_path or os.getenv(
            "TOKEN_JSON_PATH", "./config/token.json"
        )

        if not os.path.exists(token_path):
            return {}

        try:
            cfg_manager = ConfigManager(token_path)
            return cfg_manager.load()
        except Exception as e:
            logger.warning(f"加载 token.json 失败，将忽略该文件: {e}")
            return {}

    def _initialize_client(self, config: Dict[str, str]) -> None:
        """初始化客户端"""
        if self.provider == "kimi":
            self._initialize_kimi_client(config)
        elif self.provider == "doubao":
            self._initialize_doubao_client(config)
        else:
            raise ValueError("provider 必须是 'kimi' 或 'doubao'")

    def _initialize_kimi_client(self, config: Dict[str, str]) -> None:
        """初始化 Kimi 客户端"""
        moonshot_key = config["moonshot_key"]
        if not moonshot_key:
            raise ValueError(
                "Kimi 需要提供 MOONSHOT_API_KEY（可通过参数、环境变量或 token.json 提供）"
            )

        self._client = OpenAI(
            api_key=moonshot_key,
            base_url="https://api.moonshot.cn/v1"
        )
        logger.info("已初始化 Kimi 客户端")

    def _initialize_doubao_client(self, config: Dict[str, str]) -> None:
        """初始化豆包客户端"""
        ak, sk, endpoint_id = config["ak"], config["sk"], config["endpoint_id"]

        if not all([ak, sk, endpoint_id]):
            missing = []
            if not ak: missing.append("VOLC_ACCESSKEY")
            if not sk: missing.append("VOLC_SECRETKEY")
            if not endpoint_id: missing.append("ARK_ENDPOINT_ID")
            raise ValueError(
                f"豆包需提供 AK/SK/EndpointID（可通过参数、环境变量或 token.json 提供）: {', '.join(missing)}"
            )

        # TODO: 实现豆包客户端初始化
        # self._client = Ark(api_key=sk, region="cn-beijing")
        self._model = endpoint_id
        logger.info("已初始化豆包(Ark)客户端")

    def reply(
            self,
            message: str,
            system: Optional[str] = None,
            use_history: bool = True,
            stream: bool = False,
    ) -> str:
        """
        发送消息并获取回复
        :param message: 用户消息
        :param system: 系统提示词
        :param use_history: 是否使用历史记录
        :param stream: 是否使用流式输出
        :return: 助手回复
        """
        # 构建消息列表
        messages = self._build_messages(message, system, use_history)

        try:
            if self.provider == "kimi":
                return self._call_kimi(messages, stream)
            elif self.provider == "doubao":
                return self._call_doubao(messages, stream)
            else:
                raise ValueError(f"不支持的 provider: {self.provider}")
        except Exception as e:
            logger.error(f"调用 {self.provider} API 失败: {e}")
            raise

    def _build_messages(
            self, message: str, system: Optional[str], use_history: bool
    ) -> List[Dict[str, str]]:
        """构建消息列表"""
        messages = []

        if system:
            messages.append({"role": "system", "content": system})

        if use_history:
            messages.extend(self._messages)

        messages.append({"role": "user", "content": message})
        return messages

    def _call_kimi(self, messages: List[Dict[str, str]], stream: bool) -> str:
        """调用 Kimi API"""
        if not self._client:
            raise RuntimeError("Kimi 客户端未正确初始化")

        response = self._client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=self.temperature,
            stream=stream,
        )

        if stream:
            # 处理流式响应
            full_response = ""
            for chunk in response:
                if chunk.choices[0].delta.content:
                    full_response += chunk.choices[0].delta.content
            return full_response
        else:
            return response.choices[0].message.content

    def _call_doubao(self, messages: List[Dict[str, str]], stream: bool) -> str:
        """调用豆包 API"""
        # TODO: 实现豆包 API 调用
        raise NotImplementedError("豆包 API 调用暂未实现")

    def clear_history(self) -> None:
        """清空对话历史"""
        self._messages.clear()

    def get_history(self) -> List[Dict[str, str]]:
        """获取对话历史"""
        return self._messages.copy()


class WheelTimeEdit(QTimeEdit):
    """支持鼠标滚轮调整的TimeEdit"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWrapping(True)  # 允许循环滚动
        self.installEventFilter(self)

    def wheelEvent(self, event):
        """鼠标滚轮事件"""
        if not self.hasFocus():
            return

        delta = event.angleDelta().y()
        current_section = self.currentSection()

        if current_section == QTimeEdit.HourSection:
            # 调整小时
            hours = self.time().hour()
            if delta > 0:
                hours = (hours + 1) % 24
            else:
                hours = (hours - 1) % 24
            new_time = QTime(hours, self.time().minute(), self.time().second())

        elif current_section == QTimeEdit.MinuteSection:
            # 调整分钟
            minutes = self.time().minute()
            if delta > 0:
                minutes = (minutes + 1) % 60
            else:
                minutes = (minutes - 1) % 60
            new_time = QTime(self.time().hour(), minutes, self.time().second())

        elif current_section == QTimeEdit.SecondSection:
            # 调整秒
            seconds = self.time().second()
            if delta > 0:
                seconds = (seconds + 1) % 60
            else:
                seconds = (seconds - 1) % 60
            new_time = QTime(self.time().hour(), self.time().minute(), seconds)

        else:
            return

        self.setTime(new_time)
        event.accept()


class WheelSpinBox(QSpinBox):
    """支持鼠标滚轮调整的SpinBox"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.installEventFilter(self)

    def wheelEvent(self, event):
        """鼠标滚轮事件"""
        if not self.hasFocus():
            return

        delta = event.angleDelta().y()
        current_value = self.value()

        if delta > 0:
            # 向上滚动，增加值
            if current_value < 60:
                step = 1  # 小数值时步长为1
            elif current_value < 180:
                step = 5  # 中等值时步长为5
            else:
                step = 30  # 大值时步长为30
            new_value = min(self.maximum(), current_value + step)
        else:
            # 向下滚动，减少值
            if current_value <= 60:
                step = 1
            elif current_value <= 180:
                step = 5
            else:
                step = 30
            new_value = max(self.minimum(), current_value - step)

        self.setValue(new_value)
        event.accept()


class HotkeyListener(QThread):
    # 自定义信号：当按下 Esc 时触发
    hotkey_activated = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.stop_event = threading.Event()  # 控制监听循环退出

    def run(self):
        """QThread 的主执行函数"""
        def on_press(key):
            if self.stop_event.is_set():
                return False  # 停止监听器
            try:
                if key == keyboard.Key.esc:
                    self.hotkey_activated.emit()  # 安全发射信号到主线程
            except Exception as e:
                print(f"热键监听错误: {e}")

        # 启动 pynput 键盘监听（阻塞）
        with keyboard.Listener(on_press=on_press) as listener:
            listener.join()

    def stop(self):
        """安全停止监听线程"""
        self.stop_event.set()  # 触发退出
        self.quit()           # 请求线程退出
        self.wait()           # 等待线程结束

class RegionCaptureOverlay(QWidget):
    """
    区域截图覆盖层，用于选择屏幕区域
    支持多屏幕、放大镜、网格显示等功能
    """
    finished = Signal(QRect)  # 自选区确认信号
    cancelled = Signal()      # 取消操作信号

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(Qt.FramelessWindowHint |
                            Qt.WindowStaysOnTopHint |
                            Qt.Tool |
                            Qt.WindowDoesNotAcceptFocus)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setCursor(Qt.CrossCursor)

        # 多屏幕支持
        self.screens = QApplication.screens()
        self.setGeometry(self._get_combined_screen_geometry())

        # 设置窗口属性
        self.setFocusPolicy(Qt.StrongFocus)
        self.setMouseTracking(True)
        self.setFocus()

        # 选择状态
        self.start_pos = QPoint()
        self.end_pos = QPoint()
        self.is_selecting = False
        self.current_mouse_pos = QPoint()

        # 放大镜配置
        self.magnifier_size = 200
        self.magnification = 3
        self.show_magnifier = True

        # 网格和参考线
        self.show_grid = False
        self.show_crosshair = True

        # 性能优化
        self.update_timer = QTimer()
        self.update_timer.setSingleShot(True)
        self.update_timer.timeout.connect(self.update)
        self.last_mouse_pos = QPoint()

        # UI配置
        self.overlay_color = QColor(0, 0, 0, 120)
        self.selection_color = QColor(255, 0, 0, 180)
        self.info_bg_color = QColor(0, 0, 0, 200)
        self.grid_color = QColor(255, 255, 255, 80)
        self.crosshair_color = QColor(255, 255, 255, 120)

    def _get_combined_screen_geometry(self):
        """获取所有屏幕的合并几何区域"""
        combined = QRect()
        for screen in self.screens:
            combined = combined.united(screen.geometry())
        return combined

    def _get_screen_at_point(self, point: QPoint) -> QScreen:
        """获取指定点所在的屏幕"""
        for screen in self.screens:
            if screen.geometry().contains(point):
                return screen
        return QApplication.primaryScreen()

    def keyPressEvent(self, event: QKeyEvent):
        """处理键盘事件"""
        if event.key() == Qt.Key_Escape:
            self.cancel_capture()
        elif event.key() == Qt.Key_Space:
            # 空格键切换放大镜显示
            self.show_magnifier = not self.show_magnifier
            self.update()
        elif event.key() == Qt.Key_G:
            # G键切换网格显示
            self.show_grid = not self.show_grid
            self.update()
        elif event.key() == Qt.Key_C:
            # C键切换十字线显示
            self.show_crosshair = not self.show_crosshair
            self.update()
        elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            # 回车键确认当前选区
            self.confirm_selection()
        elif event.key() == Qt.Key_Plus or event.key() == Qt.Key_Equal:
            # 增加放大倍数
            self.magnification = min(8, self.magnification + 1)
            self.update()
        elif event.key() == Qt.Key_Minus:
            # 减少放大倍数
            self.magnification = max(1, self.magnification - 1)
            self.update()
        else:
            super().keyPressEvent(event)

    def mousePressEvent(self, event: QMouseEvent):
        """鼠标按下事件"""
        if event.button() == Qt.LeftButton:
            self.is_selecting = True
            self.start_pos = event.globalPosition().toPoint()
            self.end_pos = self.start_pos
            self.update()
        elif event.button() == Qt.RightButton:
            self.cancel_capture()
        elif event.button() == Qt.MiddleButton:
            # 中键重置选择
            self.start_pos = QPoint()
            self.end_pos = QPoint()
            self.is_selecting = False
            self.update()
        else:
            super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent):
        """鼠标移动事件 - 带性能优化"""
        self.current_mouse_pos = event.globalPosition().toPoint()

        # 性能优化：限制更新频率
        if (self.current_mouse_pos - self.last_mouse_pos).manhattanLength() > 2:
            self.last_mouse_pos = self.current_mouse_pos

            if self.is_selecting:
                self.end_pos = self.current_mouse_pos
                # 使用定时器延迟更新，避免过于频繁的重绘
                if not self.update_timer.isActive():
                    self.update_timer.start(16)  # ~60 FPS
            elif self.show_magnifier or self.show_crosshair:
                if not self.update_timer.isActive():
                    self.update_timer.start(16)

        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent):
        """鼠标释放事件"""
        if event.button() == Qt.LeftButton and self.is_selecting:
            self.is_selecting = False
            self.confirm_selection()
        else:
            super().mouseReleaseEvent(event)

    def confirm_selection(self):
        """确认当前选区"""
        if self.start_pos.isNull() or self.end_pos.isNull():
            self.cancel_capture()
            return

        rect = QRect(self.start_pos, self.end_pos).normalized()

        # 验证选区有效性
        if rect.width() >= 5 and rect.height() >= 5:
            # 确保选区在屏幕范围内
            screen_geometry = self._get_combined_screen_geometry()
            rect = rect.intersected(screen_geometry)

            if rect.isValid() and not rect.isEmpty():
                self.finished.emit(rect)
                self.close()
                return

        # 无效选区
        self.cancel_capture()

    def cancel_capture(self):
        """取消截图操作"""
        self.cancelled.emit()
        self.close()

    def paintEvent(self, event):
        """绘制事件"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 绘制半透明遮罩
        painter.setBrush(self.overlay_color)
        painter.setPen(Qt.NoPen)
        painter.drawRect(self.rect())

        # 绘制选区
        if not self.start_pos.isNull() and not self.end_pos.isNull():
            selected_rect = QRect(self.start_pos, self.end_pos).normalized()
            self._draw_selection(painter, selected_rect)

        # 绘制十字线（非选择状态下）
        if self.show_crosshair and not self.is_selecting:
            self._draw_crosshair(painter, self.current_mouse_pos)

        # 绘制放大镜
        if self.show_magnifier and not self.current_mouse_pos.isNull():
            self._draw_magnifier(painter, self.current_mouse_pos)

    def _draw_selection(self, painter: QPainter, rect: QRect):
        """绘制选区"""
        # 清除选区部分的遮罩
        painter.setCompositionMode(QPainter.CompositionMode_Clear)
        painter.fillRect(rect, Qt.SolidPattern)
        painter.setCompositionMode(QPainter.CompositionMode_SourceOver)

        # 绘制选区边框
        pen = QPen(self.selection_color, 2)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(rect)

        # 绘制网格
        if self.show_grid:
            self._draw_grid(painter, rect)

        # 绘制选区信息
        self._draw_selection_info(painter, rect)

        # 绘制控制点（用于调整大小）
        # self._draw_control_points(painter, rect)

    def _draw_grid(self, painter: QPainter, rect: QRect):
        """在选区内绘制网格"""
        if rect.width() < 50 or rect.height() < 50:
            return

        pen = QPen(self.grid_color, 1, Qt.DotLine)
        painter.setPen(pen)

        # 计算网格间距
        x_spacing = max(20, rect.width() // 10)
        y_spacing = max(20, rect.height() // 10)

        # 绘制垂直线
        for x in range(rect.left() + x_spacing, rect.right(), x_spacing):
            painter.drawLine(x, rect.top(), x, rect.bottom())

        # 绘制水平线
        for y in range(rect.top() + y_spacing, rect.bottom(), y_spacing):
            painter.drawLine(rect.left(), y, rect.right(), y)

    def _draw_crosshair(self, painter: QPainter, pos: QPoint):
        """绘制十字线"""
        pen = QPen(self.crosshair_color, 1, Qt.DashLine)
        painter.setPen(pen)

        # 水平线
        painter.drawLine(0, pos.y(), self.width(), pos.y())
        # 垂直线
        painter.drawLine(pos.x(), 0, pos.x(), self.height())

    def _draw_control_points(self, painter: QPainter, rect: QRect):
        """绘制选区控制点"""
        points = [
            rect.topLeft(), rect.topRight(),
            rect.bottomLeft(), rect.bottomRight(),
            QPoint(rect.center().x(), rect.top()),
            QPoint(rect.center().x(), rect.bottom()),
            QPoint(rect.left(), rect.center().y()),
            QPoint(rect.right(), rect.center().y())
        ]

        painter.setBrush(QColor(255, 255, 255, 200))
        painter.setPen(QPen(QColor(0, 0, 0, 200), 1))

        for point in points:
            painter.drawEllipse(point, 3, 3)

    def _draw_selection_info(self, painter: QPainter, rect: QRect):
        """绘制选区信息"""
        size_text = f"{rect.width()} × {rect.height()}"
        pos_text = f"({rect.x()}, {rect.y()})"
        area_text = f"Area: {rect.width() * rect.height()} px²"

        # 设置字体
        font = QFont("Arial", 10, QFont.Bold)
        painter.setFont(font)

        # 计算文本尺寸
        metrics = QFontMetrics(font)
        text_width = max(
            metrics.horizontalAdvance(size_text),
            metrics.horizontalAdvance(pos_text),
            metrics.horizontalAdvance(area_text)
        ) + 20

        text_height = 60

        # 确定信息框位置（避免超出屏幕）
        info_x = rect.right() + 10
        info_y = rect.top()

        if info_x + text_width > self.width():
            info_x = rect.left() - text_width - 10
        if info_y + text_height > self.height():
            info_y = rect.bottom() - text_height

        info_rect = QRect(info_x, info_y, text_width, text_height)

        # 绘制背景
        painter.setPen(Qt.NoPen)
        painter.setBrush(self.info_bg_color)
        painter.drawRoundedRect(info_rect, 5, 5)

        # 绘制文本
        painter.setPen(QPen(Qt.white))
        text_content = f"{size_text}\n{pos_text}\n{area_text}"
        painter.drawText(info_rect, Qt.AlignCenter, text_content)

    def _draw_magnifier(self, painter: QPainter, mouse_pos: QPoint):
        """绘制放大镜效果 - 显示在鼠标右下方，仅放大原始屏幕像素"""
        screen = self._get_screen_at_point(mouse_pos)
        if not screen:
            return

        # 计算放大镜位置（显示在鼠标右下方）
        magnifier_rect = QRect(0, 0, self.magnifier_size, self.magnifier_size)
        magnifier_rect.moveTopLeft(QPoint(mouse_pos.x() + 35, mouse_pos.y()+ 35))
        # 调整位置确保放大镜完全可见
        if magnifier_rect.right() > self.width():
            magnifier_rect.moveRight(mouse_pos.x() - 35)
        if magnifier_rect.bottom() > self.height():
            magnifier_rect.moveBottom(mouse_pos.y() - 35)
        if magnifier_rect.left() < 0:
            magnifier_rect.moveLeft(mouse_pos.x() + 35)
        if magnifier_rect.top() < 0:
            magnifier_rect.moveTop(mouse_pos.y() + 35)

        # 计算捕获区域（以鼠标位置为中心）
        capture_size = self.magnifier_size // self.magnification
        capture_rect = QRect(0, 0, capture_size, capture_size)
        capture_rect.moveCenter(mouse_pos)

        # 获取屏幕截图（仅原始屏幕内容，不包括当前绘制的放大镜）
        try:
            # 使用窗口ID为0来捕获屏幕，避免捕获到当前窗口
            screenshot = screen.grabWindow(
                0,  # 0表示捕获整个屏幕
                capture_rect.x() - screen.geometry().x(),
                capture_rect.y() - screen.geometry().y(),
                capture_rect.width(),
                capture_rect.height()
            )
        except:
            return

        # 放大绘制
        magnified = screenshot.scaled(
            self.magnifier_size,
            self.magnifier_size,
            Qt.IgnoreAspectRatio,
            Qt.SmoothTransformation
        )

        # 绘制放大镜背景和边框
        painter.setPen(QPen(QColor(255, 255, 255, 200), 2))
        painter.setBrush(QColor(0, 0, 0, 220))
        painter.drawRoundedRect(magnifier_rect, 8, 8)

        # 绘制放大内容
        painter.drawPixmap(magnifier_rect, magnified)

        # 绘制坐标和放大倍数信息
        info_text = f"({mouse_pos.x()}, {mouse_pos.y()}) {self.magnification}x"
        painter.setFont(QFont("Arial", 9))
        painter.setPen(Qt.white)

        text_rect = QRect(
            magnifier_rect.left() + 5,
            magnifier_rect.top() + 5,
            magnifier_rect.width() - 10,
            20
        )
        painter.drawText(text_rect, Qt.AlignLeft, info_text)

    def showEvent(self, event):
        """窗口显示事件"""
        super().showEvent(event)
        self.setFocus(Qt.ActiveWindowFocusReason)
        self.grabMouse()
        self.grabKeyboard()

    def hideEvent(self, event):
        """窗口隐藏事件"""
        self.releaseMouse()
        self.releaseKeyboard()
        super().hideEvent(event)

    def closeEvent(self, event):
        """窗口关闭事件"""
        self.update_timer.stop()
        super().closeEvent(event)

class StepTableHelper:
    """负责把步骤对象渲染成表格行的工具类，可放到主窗口里复用"""
    FIXED_ROW_HEIGHT = 32          # 统一行高（像素）
    ICON_SIZE = 20          # 左侧图标宽/高
    IMG_HEIGHT = 32         # 如果是图片，缩略图高度

    @staticmethod
    def widget_of(step: dict, use_color: bool = True) -> QWidget:
        """
        返回一个可直接塞进 QTableWidget 的 QWidget，
        内部 QLabel 负责显示图标/文字/图片 + 时间
        """
        t = step["type"]
        p = step["params"]
        time_str = p.get("step_time",datetime.now().strftime("%H:%M:%S"))

        # 主容器
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(4, 2, 4, 2)

        # 左侧图标或图片
        icon_label = QLabel()
        icon_label.setFixedSize(StepTableHelper.ICON_SIZE, StepTableHelper.ICON_SIZE)
        icon_label.setScaledContents(True)

        # 中间文字/图片
        content_label = QLabel()
        content_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        content_label.setStyleSheet("""color:#ffffff;
background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
border-radius:6px;
padding:2px 6px;
font-weight:bold;""")

        # 右侧时间
        time_label = QLabel(time_str)
        time_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        font = QFont()
        font.setPointSize(8)
        time_label.setFont(font)
        # 在 time_label 设置样式之前添加以下代码
        # 将时间字符串转换为颜色值
        time_obj = datetime.strptime(time_str, "%H:%M:%S")
        hour = time_obj.hour
        # minute = time_obj.minute
        # second = time_obj.second
        # 根据小时数生成低饱和度渐变色
        # 早晨(6-12): 蓝绿色调
        if 6 <= hour < 12:
            # 从浅蓝到浅绿的渐变（饱和度×1.3）
            r1, g1, b1 = 152, 196, 211  # 原 173,216,230
            r2, g2, b2 = 114, 227, 114  # 原 144,238,144

        elif 12 <= hour < 18:
            # 从浅黄到浅橙的渐变（饱和度×1.3）
            r1, g1, b1 = 255, 255, 159  # 原 255,255,224
            r2, g2, b2 = 255, 198, 137  # 原 255,218,185

        elif 18 <= hour < 21:
            # 从浅粉到浅紫的渐变（饱和度×1.3）
            r1, g1, b1 = 255, 156, 169  # 原 255,182,193
            r2, g2, b2 = 214, 214, 238  # 原 230,230,250

        else:  # 21-6 夜晚
            # 从浅蓝到浅紫的渐变（饱和度×1.3）
            r1, g1, b1 = 230, 238, 245  # 原 240,248,255
            r2, g2, b2 = 214, 214, 238  # 原 230,230,250
        if use_color:
            time_label.setStyleSheet(f"""color:#ffffff;
            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 rgb({r1},{g1},{b1}),stop:1 rgb({r2},{g2},{b2}));
            border-radius:10px;
            padding:2px 6px;
            font-weight:bold;""")
        else:
            time_label.setStyleSheet("""color:#ffffff;
            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
            border-radius:6px;
            padding:2px 6px;
            font-weight:bold;""")

        # 根据类型生成内容
        if t == "鼠标点击":
            use_image = p.get("use_image", True)
            use_coordinates = p.get("use_coordinates", False)

            if use_image:
                img_path = p.get("image_path", "")
                click_type = p.get("click_type", "左键单击")
                if os.path.isfile(img_path):
                    pm = QPixmap(img_path).scaledToHeight(StepTableHelper.IMG_HEIGHT, Qt.SmoothTransformation)
                    icon_label.setPixmap(pm)
                else:
                    icon_label.setText("🖼️")
                content_label.setText(f"{click_type} · 图片")

            elif use_coordinates:
                x_coord = p.get("x_coordinate", 0)
                y_coord = p.get("y_coordinate", 0)
                click_type = p.get("click_type", "左键单击")
                icon_label.setText("📍")
                content_label.setText(f"{click_type}\n坐标({x_coord},{y_coord})")

            else:
                icon_label.setText("❓")
                content_label.setText("未设置模式")
            click_type = p.get("click_type", "左键单击")
            # 为不同点击类型设置不同的低饱和度渐变背景
            if use_color:
                if click_type == "左键单击":
                    content_label.setStyleSheet("""color:#ffffff;
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a8c0ff,stop:1 #a8c0ff);
                        border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif click_type == "左键双击":
                    content_label.setStyleSheet("""color:#ffffff;
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d4fc79,stop:1 #96e6a1);
                        border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif click_type == "右键单击":
                    content_label.setStyleSheet("""color:#ffffff;
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #f6d365,stop:1 #fda085);
                        border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif click_type == "中键单击":
                    content_label.setStyleSheet("""color:#ffffff;
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #84fab0,stop:1 #8fd3f4);
                        border-radius:6px;padding:2px 6px;font-weight:bold;""")
            else:
                # 默认样式（如果出现其他点击类型）
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")

        elif t == "文本输入":
            txt = p.get("text", "")
            if txt:
                txt = txt[:10] + "…" if len(txt) > 10 else txt
                content_label.setText(txt)
            else:
                mode = p.get("mode", "顺序")
                file = os.path.basename(p.get("excel_path", ""))
                content_label.setText(f"{mode}·{file}")
            icon_label.setText("⌨")
            # 设置低饱和度渐变背景
            if use_color:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d4fc79,stop:1 #96e6a1);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")
            else:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")


        elif t == "等待":
            content_label.setText(f"{p.get('seconds', 0)}s")
            icon_label.setText("⏱")
            # 设置低饱和度渐变背景
            if use_color:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #f6d365,stop:1 #fda085);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")
            else:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")


        elif t == "截图":
            save_path = p.get("save_path", "")
            if os.path.isfile(save_path):
                pm = QPixmap(save_path).scaledToHeight(StepTableHelper.IMG_HEIGHT, Qt.SmoothTransformation)
                content_label.setPixmap(pm)
            else:
                content_label.setText(os.path.basename(save_path))
            icon_label.setText("📸")

        elif t == "鼠标滚轮":
            dire = p.get("direction", "向下")
            clicks = p.get("clicks", 3)
            content_label.setText(f"{dire}{clicks}格")
            icon_label.setText("⚙")
            # 设置低饱和度渐变背景
            if use_color:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a6c0fe,stop:1 #f68084);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")
            else:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")

        elif t == "键盘热键":
            hotkey = p.get("hotkey", "ctrl+c").upper()
            delay = p.get("delay_ms", 100)
            content_label.setText(f"{hotkey}")
            time_label.setText(f"{delay} ms")
            icon_label.setText("⌨")
            # 设置低饱和度渐变背景
            if use_color:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d299c2,stop:1 #fef9d7);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")
            else:
                content_label.setStyleSheet("""color:#ffffff;
                    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                    border-radius:6px;padding:2px 6px;font-weight:bold;""")
        elif t == "拖拽":
            use_image = p.get("use_image", True)
            # 清除可能存在的旧图片
            icon_label.setText("")
            icon_label.setPixmap(QPixmap())

            if use_image:
                img_path = p.get("image_path", "")
                # 根据拖拽方向确定显示文本
                dx, dy = p.get("drag_x", 0), p.get("drag_y", 100)
                if dx == 0 and dy > 0:
                    content_label.setText("↓下拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a8c0ff,stop:1 #a8c0ff);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif dx == 0 and dy < 0:
                    content_label.setText("↑上拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d4fc79,stop:1 #96e6a1);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif dx > 0 and dy == 0:
                    content_label.setText("→右拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #f6d365,stop:1 #fda085);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif dx < 0 and dy == 0:
                    content_label.setText("←左拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #84fab0,stop:1 #8fd3f4);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                else:
                    content_label.setText(f"图像拖拽 ({dx},{dy})")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #fbc2eb,stop:1 #a6c1ee);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                if not use_color:
                    content_label.setStyleSheet("""color:#ffffff;
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                        border-radius:6px;padding:2px 6px;font-weight:bold;""")


                # 在icon_label中显示图片缩略图
                if os.path.isfile(img_path):
                    pm = QPixmap(img_path).scaledToHeight(StepTableHelper.ICON_SIZE, Qt.SmoothTransformation)
                    icon_label.setPixmap(pm)
                else:
                    icon_label.setText("✋")  # 图片不存在时显示手型图标
            else:
                sx, sy = p.get("start_x", 0), p.get("start_y", 0)
                ex, ey = p.get("end_x", 0), p.get("end_y", 0)
                # 根据坐标变化显示箭头
                dx, dy = ex - sx, ey - sy
                if dx == 0 and dy > 0:
                    content_label.setText("↓下拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a8c0ff,stop:1 #a8c0ff);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif dx == 0 and dy < 0:
                    content_label.setText("↑上拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d4fc79,stop:1 #96e6a1);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif dx > 0 and dy == 0:
                    content_label.setText("→右拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #f6d365,stop:1 #fda085);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                elif dx < 0 and dy == 0:
                    content_label.setText("←左拉")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #84fab0,stop:1 #8fd3f4);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                else:
                    content_label.setText(f"坐标拖拽 ({sx},{sy})→({ex},{ey})")
                    content_label.setStyleSheet("""color:#ffffff;
                            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #fbc2eb,stop:1 #a6c1ee);
                            border-radius:6px;padding:2px 6px;font-weight:bold;""")
                if not use_color:
                    content_label.setStyleSheet("""color:#ffffff;
                        background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
                        border-radius:6px;padding:2px 6px;font-weight:bold;""")
                icon_label.setText("✋")
        elif t == "自动回复":
            provider = p.get("provider", "kimi")
            content_label.setText(f"{provider}")
            icon_label.setText("🤖")

            # 设置样式
            if use_color:
                content_label.setStyleSheet("""color:#ffffff;
            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a6c0fe,stop:1 #f68084);
            border-radius:6px;
            padding:2px 6px;
            font-weight:bold;""")
            else:
                content_label.setStyleSheet("""color:#ffffff;
            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
            border-radius:6px;
            padding:2px 6px;
            font-weight:bold;""")
        else:
            content_label.setText(t)
            icon_label.setText("?")

        # 加入布局
        layout.addWidget(icon_label)
        layout.addWidget(content_label, 1)   # 伸缩
        layout.addWidget(time_label)

        return container


    @staticmethod
    def thumb_widget(img_path: str, row_height: int) -> QWidget:
        """返回一个已设置好缩略图的 QLabel，高度=row_height，宽度自适应"""
        label = QLabel()
        label.setScaledContents(True)
        label.setAlignment(Qt.AlignCenter)

        # 读图并缩放到行高
        pixmap = QPixmap(img_path)
        if not pixmap.isNull():
            pixmap = pixmap.scaledToHeight(row_height, Qt.SmoothTransformation)
        label.setPixmap(pixmap)

        # 用 QWidget 包一层，方便后续扩展
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.addWidget(label)
        return w

    @staticmethod
    def type_widget(step_type: str, use_color: bool = True) -> QWidget:
        """
        创建一个用于显示步骤类型的QWidget容器，可以直接添加到表格中

        Args:
            step_type: 步骤类型
            use_color: 是否使用彩色样式，False时使用黑灰色调样式

        Returns:
            QWidget: 包含图标和类型标签的容器
        """
        # 创建主容器
        container = QWidget()
        layout = QHBoxLayout(container)
        layout.setContentsMargins(4, 2, 4, 2)
        layout.setAlignment(Qt.AlignCenter)

        # 创建图标标签
        icon_label = QLabel()
        icon_label.setAlignment(Qt.AlignCenter)
        icon_label.setStyleSheet("font-size: 14px; margin-right: 5px;")

        # 根据步骤类型设置对应图标
        icons = {
            "鼠标点击": "🖱️",
            "文本输入": "⌨️",
            "等待": "⏱️",
            "截图": "📸",
            "拖拽": "✋",
            "鼠标滚轮": "🖱️",  # 使用相同图标但可以区分
            "键盘热键": "⌨️",
            "自动回复": "🤖"
        }

        icon_text = icons.get(step_type, "❓")  # 默认问号图标
        icon_label.setText(icon_text)

        # 创建类型标签
        type_label = QLabel(step_type)
        type_label.setAlignment(Qt.AlignCenter)

        # 设置样式
        if not use_color:
            # 统一使用黑灰色调
            type_label.setStyleSheet("""color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #555555,stop:0.5 #777777,stop:1 #999999);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""")
        else:
            # 根据不同类型返回不同颜色样式
            styles = {
                "鼠标点击": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a8c0ff,stop:1 #a8c0ff);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",

                "文本输入": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d4fc79,stop:1 #96e6a1);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",

                "等待": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #f6d365,stop:1 #fda085);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",

                "截图": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #84fab0,stop:1 #8fd3f4);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",

                "拖拽": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #fbc2eb,stop:1 #a6c1ee);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",

                "鼠标滚轮": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a6c0fe,stop:1 #f68084);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",

                "键盘热键": """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #d299c2,stop:1 #fef9d7);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""",
                "自动回复": """color:#ffffff;
            background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #a6c0fe,stop:1 #f68084);
            border-radius:6px;
            padding:2px 6px;
            font-weight:bold;"""  # 添加这一行
            }

            # 设置对应样式或默认样式
            style = styles.get(step_type, """color:#ffffff;
    background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #e5e5e5,stop:0.5 #bdbdbd,stop:1 #9e9e9e);
    border-radius:6px;
    padding:2px 6px;
    font-weight:bold;""")
            type_label.setStyleSheet(style)

        # 添加到布局
        layout.addWidget(icon_label)
        layout.addWidget(type_label)

        return container


class ATIcon:
    """
    为「Auto Tool」桌面自动办公软件生成一枚
    64×64 带毛玻璃效果、渐变背景的「AT」图标。
    """
    SIZE = 64
    _cache = {}          # 缓存，避免重复渲染

    @classmethod
    def pixmap(cls, size=SIZE) -> QPixmap:
        """返回渲染好的 QPixmap，可自由缩放"""
        if size in cls._cache:
            return cls._cache[size]

        px = QPixmap(size, size)
        px.fill(Qt.transparent)

        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)

        # 1. 圆角矩形背景 -------------------------------------------------
        rect = QRectF(0, 0, size, size)
        radius = size * 0.18
        path = QPainterPath()
        path.addRoundedRect(rect, radius, radius)

        # 2. 渐变填充 ------------------------------------------------------
        g = QLinearGradient(QPointF(0, 0), QPointF(size, size))
        g.setColorAt(0.0, QColor("#6A11CB"))   # 紫
        g.setColorAt(1.0, QColor("#2575FC"))   # 蓝
        p.fillPath(path, QBrush(g))

        # 3. 毛玻璃：一层极低不透明度白色蒙版 -------------------------------
        blur_layer = QPainterPath()
        blur_layer.addRoundedRect(rect, radius, radius)
        p.fillPath(blur_layer, QColor(255, 255, 255, 35))

        # 4. 字母 “AT” ----------------------------------------------------
        font = QFont("Segoe UI", size * 0.32, QFont.Bold)
        p.setFont(font)
        p.setPen(QPen(Qt.white))
        p.drawText(rect, Qt.AlignCenter, "AT")

        p.end()
        cls._cache[size] = px
        return px

    @classmethod
    def icon(cls, size=SIZE) -> QIcon:
        """直接拿到 QIcon，可设给窗口、托盘、按钮等"""
        return QIcon(cls.pixmap(size))

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("AboutDialog")
        self.setWindowTitle("关于")
        self.setModal(True)
        self.resize(480, 520)

        # 根布局
        root = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        root.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        lay = QVBoxLayout(content)
        lay.setAlignment(Qt.AlignTop)

        # 1. 标题
        title = QLabel("自动化任务管理器")
        title.setObjectName("titleLabel")
        title.setAlignment(Qt.AlignCenter)
        lay.addWidget(title)
        # 2. 版本 + 作者 + 头像
        author_layout = QHBoxLayout()
        author_layout.setSpacing(12)

        # 头像
        self.avatar = QLabel()
        self.avatar.setFixedSize(64, 64)
        self.avatar.setObjectName("avatarLabel")
        self.load_avatar()

        # 作者信息
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignRight)
        form.addRow("版　本：", QLabel("1.0.0"))
        form.addRow("作　者：", QLabel("B_arbarian from UESTC"))
        author_layout.addWidget(self.avatar)
        author_layout.addLayout(form)
        author_layout.addStretch()
        lay.addLayout(author_layout)
        # 3. 联系方式（带超链接）
        link_lbl = QLabel(
            'B站主页：<a href="https://space.bilibili.com/521967044">'
            '<span style="color:#409EFF;">点击访问</span></a><br>'
            '邮　　箱：<a href="mailto:264214429@qq.com">'
            '<span style="color:#409EFF;">264214429@qq.com</span></a>'
        )
        link_lbl.setObjectName("linkLabel")
        link_lbl.setOpenExternalLinks(True)
        link_lbl.setTextInteractionFlags(Qt.TextBrowserInteraction)
        lay.addWidget(link_lbl, alignment=Qt.AlignCenter)

        # 4. 简介
        intro = QTextEdit()
        intro.setObjectName("introText")
        intro.setMaximumHeight(180)
        intro.setPlainText(
            "自动化任务管理器是一款强大的桌面自动化工具，可以帮助您自动化执行重复的计算机操作，提高工作效率。\n\n"
            "主要功能：\n"
            "• 基于图像识别的鼠标操作\n"
            "• 文本输入自动化\n"
            "• 定时任务执行\n"
            "• 详细执行日志记录\n\n"
            "感谢使用本软件！如有任何问题或建议，请通过上述联系方式与我们联系。"
        )
        lay.addWidget(intro)
        # 5. 打赏二维码
        qr_lay = QHBoxLayout()
        qr_lay.setSpacing(16)
        qr_lay.addStretch()

        self.wx_qr = QLabel()
        self.wx_qr.setObjectName("qrLabel")
        self.wx_qr.setFixedSize(160, 160)
        self.load_qr(self.wx_qr, "img/donate.png", "微信赞赏")

        self.zfb_qr = QLabel()
        self.zfb_qr.setObjectName("qrLabel")
        self.zfb_qr.setFixedSize(160, 160)
        self.load_qr(self.zfb_qr, "img/zhifubao.jpg", "支付宝打赏")

        qr_lay.addWidget(self.wx_qr)
        qr_lay.addWidget(self.zfb_qr)
        qr_lay.addStretch()
        lay.addLayout(qr_lay)

        # 6. 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok)
        btn_box.accepted.connect(self.accept)
        root.addWidget(btn_box)

        # 7. 加载样式
        self.load_qss()

    # ---------- 私有方法 ----------
    def load_avatar(self):
        avatar_path = "img/avatar.jpg"
        pixmap = QPixmap(avatar_path)
        if pixmap.isNull():
            self.avatar.setText("头像")
            return

        size = self.avatar.width()
        rounded = QPixmap(size, size)
        rounded.fill(Qt.transparent)

        painter = QPainter(rounded)
        painter.setRenderHint(QPainter.Antialiasing)
        path = QPainterPath()
        path.addRoundedRect(0, 0, size, size, size//2, size//2)
        painter.setClipPath(path)
        painter.drawPixmap(0, 0, size, size, pixmap)
        painter.end()

        self.avatar.setPixmap(rounded)


    def load_qr(self, label: QLabel, path: str, alt: str):
        path = resource_path(path)
        pixmap = QPixmap(path)
        if pixmap.isNull():
            label.setText(f"{alt}\n加载失败")
            label.setAlignment(Qt.AlignCenter)
            return
        label.setPixmap(pixmap.scaled(label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
    def load_qss(self):
        qss_path = Path(__file__).resolve().parent.parent / "css" / "about_style.qss"
        if qss_path.exists():
            with open(qss_path, 'r', encoding='utf-8') as f:
                self.setStyleSheet(f.read())


class TaskRunner(QObject):
    task_completed = Signal(str, bool, str)
    task_progress = Signal(str, int, int)
    task_stopped = Signal(str)
    log_message = Signal(str, str)  # 新增日志信号

    def __init__(self, task_name, steps,auto_skip_image_timeout=False,timeout=10,instant_click=False,move_duration=0.1,parent=None):
        super().__init__()
        self.task_name = task_name
        self.steps = steps
        self.is_running = False
        self.current_step = 0
        self.repeat_count = 0
        self.max_repeat = 1  # 默认执行1次
        self.repeat_interval = 0

        self.auto_skip_image_timeout = auto_skip_image_timeout
        self.timeout = timeout  # 用户设置的超时时间

        self.instant_click = instant_click        # 是否跳过移动动画
        self.default_move_duration = move_duration  # 全局移动动画时长

        self._excel_cycle = None
        self._excel_cache = {}   # 路径->(wb, ws, rows)

        self.parent =  parent

    def set_repeat_interval(self, interval_minutes):
        """
        设置重复间隔时间

        Args:
            interval_minutes (int): 间隔时间（分钟）
        """
        self.repeat_interval = interval_minutes
    def set_repeat_count(self, count):
        self.max_repeat = count

    def execute_mouse_click(self, params):
        """
        执行鼠标点击操作
        支持图片识别点击和坐标直接点击两种模式
        """
        use_image = params.get("use_image", True)
        use_coordinates = params.get("use_coordinates", False)

        # 检查参数有效性
        if use_image and use_coordinates:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, "⚠️ 图片和坐标模式不能同时启用，跳过此步骤")
                return
            else:
                raise ValueError("图片和坐标模式不能同时启用")

        if not use_image and not use_coordinates:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, "⚠️ 未启用图片模式也未启用坐标模式，跳过此步骤")
                return
            else:
                raise ValueError("必须启用图片模式或坐标模式")

        # 获取通用参数
        click_type = params.get("click_type", "左键单击")
        offset_x = params.get("offset_x", 0)
        offset_y = params.get("offset_y", 0)
        move_duration = params.get("move_duration", self.default_move_duration)

        # 点击类型映射
        click_map = {
            "左键单击": pyautogui.click,
            "左键双击": pyautogui.doubleClick,
            "右键单击": pyautogui.rightClick,
            "中键单击": pyautogui.middleClick,
        }

        if click_type not in click_map:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, f"⚠️ 不支持的点击类型: {click_type}，跳过")
                return
            else:
                raise ValueError(f"不支持的 click_type: {click_type}")

        # 模式1: 使用坐标直接点击
        if use_coordinates:
            x_coordinate = params.get("x_coordinate", 0)
            y_coordinate = params.get("y_coordinate", 0)

            if x_coordinate == 0 and y_coordinate == 0:
                if self.auto_skip_image_timeout:
                    self.log_message.emit(self.task_name, "⚠️ 坐标不能都为0，跳过此步骤")
                    return
                else:
                    raise ValueError("坐标不能都为0")

            target_x = x_coordinate + offset_x
            target_y = y_coordinate + offset_y

            self.log_message.emit(self.task_name,
                                  f"📌 使用坐标模式: ({x_coordinate}, {y_coordinate}) + 偏移({offset_x}, {offset_y}) = 目标({target_x}, {target_y})")

            # 移动鼠标
            if not self.instant_click:
                try:
                    pyautogui.moveTo(target_x, target_y, duration=move_duration)
                except Exception as e:
                    if self.auto_skip_image_timeout:
                        self.log_message.emit(self.task_name, f"⚠️ 鼠标移动失败，跳过: {e}")
                        return
                    raise
            else:
                pyautogui.moveTo(target_x, target_y, duration=0)  # 瞬移

            # 执行点击
            click_map[click_type](target_x, target_y)
            self.log_message.emit(self.task_name, f"✅ 已完成坐标 {click_type} 操作")
            return

        # 模式2: 使用图片识别点击
        image_path = params.get("image_path", "")
        scan_direction = params.get("scan_direction", "默认")
        confidence = params.get("confidence", 0.8)
        timeout = params.get("timeout", self.timeout)

        if not image_path:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, "⚠️ 图片路径为空，跳过此步骤")
                return
            else:
                raise ValueError("image_path 不能为空")

        if not os.path.exists(image_path):
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name, f"⚠️ 图片文件不存在: {image_path}，跳过此步骤")
                return
            else:
                raise FileNotFoundError(f"图片文件不存在: {image_path}")

        self.log_message.emit(self.task_name, f"🔍 开始定位图片: {os.path.basename(image_path)}")
        self.log_message.emit(self.task_name, f"📊 扫描方向: {scan_direction}, 置信度: {confidence}, 超时: {timeout}s")

        def find_image_center():
            """默认查找图片中心"""
            start = time.time()
            while True:
                pos = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
                if pos:
                    return pos
                if time.time() - start > timeout:
                    return None
                time.sleep(0.2)

        def find_image_center_with_direction():
            """
            按指定方向返回第一个匹配图的中心坐标。
            direction: "从左到右" | "从右到左" | "从上到下" | "从下到上"
            """
            start = time.time()
            while True:
                # 1. 拿到所有匹配框
                boxes = list(pyautogui.locateAllOnScreen(image_path, confidence=confidence))
                if boxes:
                    # 2. 按方向排序
                    if scan_direction == "从左到右":
                        boxes.sort(key=lambda b: b.left)  # left 升序
                    elif scan_direction == "从右到左":
                        boxes.sort(key=lambda b: -(b.left + b.width))  # 最右在前
                    elif scan_direction == "从上到下":
                        boxes.sort(key=lambda b: b.top)  # top 升序
                    elif scan_direction == "从下到上":
                        boxes.sort(key=lambda b: -(b.top + b.height))  # 最下在前
                    else:
                        # 防呆，回到默认（最左上）
                        boxes.sort(key=lambda b: (b.top, b.left))

                    # 3. 取第一个框的中心
                    target = boxes[0]
                    x, y = pyautogui.center(target)
                    return (x, y)
                # 4. 超时判定
                if time.time() - start > timeout:
                    return None
                time.sleep(0.2)

        # 执行图片查找
        if scan_direction == "默认":
            center = find_image_center()
        else:
            center = find_image_center_with_direction()

        if center is None:
            if self.auto_skip_image_timeout:
                self.log_message.emit(self.task_name,
                                      f"⚠️ 在 {timeout}s 内未找到图片: {os.path.basename(image_path)}，自动跳过")
                return  # ✅ 跳过，不抛异常
            else:
                raise RuntimeError(f"在 {timeout}s 内未找到图片: {image_path}")

        # 计算目标坐标（考虑偏移）
        if scan_direction == "默认":
            target_x = center.x + offset_x
            target_y = center.y + offset_y
        else:
            target_x = center[0] + offset_x
            target_y = center[1] + offset_y

        self.log_message.emit(self.task_name,
                              f"🎯 找到图片位置: ({center.x if scan_direction == '默认' else center[0]}, {center.y if scan_direction == '默认' else center[1]}) + 偏移({offset_x}, {offset_y}) = 目标({target_x}, {target_y})")

        # 移动鼠标
        if not self.instant_click:
            try:
                pyautogui.moveTo(target_x, target_y, duration=move_duration)
            except Exception as e:
                if self.auto_skip_image_timeout:
                    self.log_message.emit(self.task_name, f"⚠️ 鼠标移动失败，跳过: {e}")
                    return
                raise
        else:
            pyautogui.moveTo(target_x, target_y, duration=0)  # 瞬移

        # 执行点击
        click_map[click_type](target_x, target_y)
        self.log_message.emit(self.task_name, f"✅ 已完成图片 {click_type} 操作")
    def run(self):
        self.is_running = True
        self.current_step = 0
        total_steps = len(self.steps)
        self.repeat_count = 0

        self.log_message.emit(self.task_name, f"🚀 开始执行任务: {self.task_name}, 共 {total_steps} 个步骤")

        try:
            while self.repeat_count < self.max_repeat and self.is_running:
                self.repeat_count += 1
                if self.max_repeat > 1:
                    self.log_message.emit(self.task_name, f"🔄 第 {self.repeat_count}/{self.max_repeat} 次执行")
                    self.parent.statusBar().showMessage(f"【{self.task_name}】第 {self.repeat_count}/{self.max_repeat} 次执行")
                for i, step in enumerate(self.steps):
                    if not self.is_running:
                        self.log_message.emit(self.task_name, "⏹️ 任务被中断")
                        break

                    self.current_step = i
                    self.task_progress.emit(self.task_name, i + 1, total_steps)

                    # 执行步骤
                    step_type = step.get("type", "")
                    params = step.get("params", {})
                    delay = step.get("delay", 0)

                    # 简化日志显示
                    if step_type == "鼠标点击":
                        use_image = params.get("use_image", True)
                        use_coordinates = params.get("use_coordinates", False)

                        self.log_message.emit(self.task_name, f"📝 执行步骤 {i + 1}/{total_steps}: {step_type}")

                        if use_image:
                            image_name = os.path.basename(params.get("image_path", "")) if params.get(
                                "image_path") else "未设置"
                            click_type = params.get("click_type", "左键单击")
                            scan_direction = params.get("scan_direction", "默认")
                            offset_x = params.get("offset_x", 0)
                            offset_y = params.get("offset_y", 0)

                            log_text = f"🖼️ 图片模式: {image_name}, 点击: {click_type}, 方向: {scan_direction}"
                            if offset_x != 0 or offset_y != 0:
                                log_text += f", 偏移: ({offset_x}, {offset_y})"
                            self.log_message.emit(self.task_name, log_text)

                        elif use_coordinates:
                            x_coord = params.get("x_coordinate", 0)
                            y_coord = params.get("y_coordinate", 0)
                            click_type = params.get("click_type", "左键单击")
                            offset_x = params.get("offset_x", 0)
                            offset_y = params.get("offset_y", 0)

                            log_text = f"📍 坐标模式: ({x_coord}, {y_coord}), 点击: {click_type}"
                            if offset_x != 0 or offset_y != 0:
                                log_text += f", 偏移: ({offset_x}, {offset_y})"
                            self.log_message.emit(self.task_name, log_text)

                        else:
                            self.log_message.emit(self.task_name, "⚠️ 未启用图片或坐标模式")
                    else:
                        self.log_message.emit(self.task_name, f"📝 执行步骤 {i + 1}/{total_steps}: {step_type}")
                        self.log_message.emit(self.task_name, f"⚙️ 参数: {json.dumps(params, ensure_ascii=False)}")
                    if step_type == "鼠标点击":
                        self.execute_mouse_click(params)
                    elif step_type == "文本输入":
                        self.execute_keyboard_input(params)
                    elif step_type == "等待":
                        self.execute_wait(params)
                    elif step_type == "截图":
                        self.execute_screenshot(params)
                    elif step_type == "拖拽":
                        self.execute_drag(params)
                    elif step_type == "鼠标滚轮":
                        self.execute_mouse_scroll(params)
                    elif step_type == "键盘热键":
                        self.execute_hotkey(params)
                    elif step_type == "自动回复":
                        self.execute_ai_reply(params)
                    else:
                        self.log_message.emit(self.task_name, f"⚠️ 未知步骤类型: {step_type}")

                    # 步骤间延时
                    if delay > 0:
                        self.log_message.emit(self.task_name, f"⏱️ 步骤延时: {delay}秒")
                        time.sleep(delay)

                # 检查是否需要等待下次重复执行
                if self.repeat_count < self.max_repeat and self.is_running and self.repeat_interval > 0:
                    wait_seconds = self.repeat_interval * 60  # 转换为秒
                    countdown_start = wait_seconds - 10  # 提前10秒开始倒计时
                    self.log_message.emit(self.task_name, f"⏳ 间隔等待: {self.repeat_interval}分钟")
                    self.parent.statusBar().showMessage(
                        f"【{self.task_name}】⏳ 间隔等待: {self.repeat_interval}分钟")
                    # 分段等待，每秒检查一次是否停止
                    for _ in range(int(countdown_start)):
                        if not self.is_running:
                            self.log_message.emit(self.task_name, "⏹️ 任务被中断")
                            break
                        time.sleep(1)
                    # 开始10秒倒计时
                    countdown_seconds = 10
                    while countdown_seconds > 0 and self.is_running:
                        current_time = time.strftime('%H:%M:%S')  # 获取当前时间
                        self.parent.statusBar().showMessage(
                            f"[{current_time}]【{self.task_name}】⏳ 倒计时: {countdown_seconds} 秒")
                        time.sleep(1)
                        countdown_seconds -= 1
                if not self.is_running:
                    break
            success = self.is_running
            message = "✅ 任务完成" if success else "⏹️ 任务被中断"
            self.log_message.emit(self.task_name, message)
            self.parent.statusBar().showMessage(message)
            self.task_completed.emit(self.task_name, success, message)
        except Exception as e:
            error_msg = f"❌ 任务执行出错: {str(e)}"
            self.log_message.emit(self.task_name, error_msg)
            self.task_completed.emit(self.task_name, False, error_msg)
        finally:
            self.is_running = False

    def stop(self):
        self.log_message.emit(self.task_name, "⏹️ 停止任务")
        self.is_running = False
        self.task_stopped.emit(self.task_name)
    def chinese_qixi(self,year: int) -> date:
        """
        计算指定年份的七夕节（农历七月初七）的公历日期
        使用近似算法，误差在±1天内

        Args:
            year: 要计算的年份

        Returns:
            该年份七夕节的公历日期
        """
        # 扩展的年份对照表（2000-2030年）
        table = {
            2000: date(2000, 8, 6), 2001: date(2001, 8, 25), 2002: date(2002, 8, 15),
            2003: date(2003, 8, 4), 2004: date(2004, 8, 22), 2005: date(2005, 8, 11),
            2006: date(2006, 7, 31), 2007: date(2007, 8, 19), 2008: date(2008, 8, 7),
            2009: date(2009, 8, 26), 2010: date(2010, 8, 16), 2011: date(2011, 8, 6),
            2012: date(2012, 8, 23), 2013: date(2013, 8, 13), 2014: date(2014, 8, 2),
            2015: date(2015, 8, 20), 2016: date(2016, 8, 9), 2017: date(2017, 8, 28),
            2018: date(2018, 8, 17), 2019: date(2019, 8, 7), 2020: date(2020, 8, 25),
            2021: date(2021, 8, 14), 2022: date(2022, 8, 4), 2023: date(2023, 8, 22),
            2024: date(2024, 8, 10), 2025: date(2025, 8, 1), 2026: date(2026, 8, 19),
            2027: date(2027, 8, 8), 2028: date(2028, 7, 28), 2029: date(2029, 8, 16),
            2030: date(2030, 8, 5)
        }

        # 如果在已知年份范围内，直接返回表中日期
        if year in table:
            return table[year]

        # 对于表外的年份，使用近似算法计算
        # 基础年份选择2023年，七夕日期为8月22日
        base_year = 2023
        base_date = date(base_year, 8, 22)

        # 计算与基础年份的差异（考虑农历年的平均长度）
        year_diff = year - base_year
        # 农历年平均长度约为29.53天×12个月 = 354.36天
        days_diff = round(year_diff * 354.36 - year_diff * 365.25)

        # 计算预估日期
        estimated_date = base_date + timedelta(days=days_diff)

        # 调整到8月附近（七夕通常在7月底到8月底之间）
        if estimated_date.month < 7:
            estimated_date += timedelta(days=30)
        elif estimated_date.month > 9:
            estimated_date -= timedelta(days=30)

        return estimated_date

    # def execute_mouse_click(self, params):
    #     AutoClicker().execute_mouse_click(params)
    #     self.log_message.emit(self.task_name, "🖱️ 鼠标点击操作完成")

    def execute_mouse_scroll(self, params):
        direction = params.get("direction", "向下滚动")
        clicks = params.get("clicks", 3)

        self.log_message.emit(self.task_name,
                              f"🖱 鼠标滚轮 {direction} {clicks} 格（当前位置）")

        try:
            scroll_amount = clicks * 120 if direction == "向下滚动" else -clicks * 120
            pyautogui.scroll(scroll_amount)
            self.log_message.emit(self.task_name, "✅ 滚轮完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 滚轮出错: {str(e)}")
            raise

    def execute_ai_reply(self, params):
        try:
            # 获取剪贴板内容作为消息
            if params.get("use_clipboard", False):
                clipboard_content = pyperclip.paste()
            else:
                clipboard_content = '用户未输入消息 根据系统提示词输出'
            # 获取参数
            provider = params.get("provider", "kimi")
            system_prompt = params.get("system_prompt", "")
            use_history = params.get("use_history", True)
            stream = params.get("stream", False)

            # 初始化 ChatBot
            bot = ChatBot(
                provider=provider,
                token_json_path="./config/token.json"
            )

            # 发送消息并获取回复
            reply = bot.reply(
                message=clipboard_content,
                system=system_prompt,
                use_history=use_history,
                stream=stream
            )

            # 将回复复制到剪贴板
            pyperclip.copy(reply)

            self.log_message.emit(self.task_name, f"✅ AI 回复成功: {reply}")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ AI 回复出错: {str(e)}")
            raise

    def execute_hotkey(self, params):
        hotkey = params.get("hotkey", "")
        delay = params.get("delay_ms", 100)

        if not hotkey:
            self.log_message.emit(self.task_name, "⚠️ 未设置热键")
            return

        self.log_message.emit(self.task_name, f"⌨ 热键 {hotkey} 执行")

        try:
            # 解析热键字符串
            keys = hotkey.lower().split("+")

            # 转换为pyautogui可识别的键名
            pyautogui_keys = []
            for key in keys:
                # 处理特殊键名映射
                key_map = {
                    "ctrl": "ctrl",
                    "alt": "alt",
                    "shift": "shift",
                    "win": "win",
                    "cmd": "cmd",
                    "enter": "enter",
                    "return": "enter",
                    "space": "space",
                    "tab": "tab",
                    "esc": "esc",
                    "escape": "esc",
                    "backspace": "backspace",
                    "delete": "delete",
                    "insert": "insert",
                    "home": "home",
                    "end": "end",
                    "pageup": "pageup",
                    "pagedown": "pagedown",
                    "up": "up",
                    "down": "down",
                    "left": "left",
                    "right": "right",
                    "capslock": "capslock",
                    "numlock": "numlock",
                    "scrolllock": "scrolllock"
                }

                if key in key_map:
                    pyautogui_keys.append(key_map[key])
                else:
                    pyautogui_keys.append(key)

            # 执行热键
            if len(pyautogui_keys) == 1:
                pyautogui.press(pyautogui_keys[0])
            else:
                pyautogui.hotkey(*pyautogui_keys)

            if delay > 0:
                time.sleep(delay / 1000.0)
            self.log_message.emit(self.task_name, "✅ 热键完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 热键出错: {str(e)}")
            raise

    def execute_keyboard_input(self, params):
        from datetime import datetime, date, time
        # 1. 纯文本优先
        text = params.get("text", "").strip()
        if not text or '未来也要一起闪耀' in text:
            # 2. 动态纪念日文案
            love_str = params.get("love_date")
            if love_str:
                love_dt = datetime.fromisoformat(love_str)
                today = date.today()
                today_1314 = datetime.combine(today, time(13, 14))

                delta = today_1314 - love_dt
                days, sec = delta.days, delta.seconds
                hours, rem = divmod(sec, 3600)
                minutes, secs = divmod(rem, 60)
                duration = f"{days}天{hours}时{minutes}分{secs}秒"

                year_start = datetime(today.year, 1, 1, 13, 14)
                count = (today_1314 - year_start).days + 1

                # 特殊节日
                is_xmas = (love_dt.month, love_dt.day) == (12, 25)
                special = ""
                if today == date(today.year, 12, 25):
                    special = "\n圣诞快乐，Merry Christmas！"
                elif today == date(today.year, 2, 14):
                    special = "\n情人节快乐！"
                elif today == self.chinese_qixi(today.year):
                    special = "\n七夕快乐，鹊桥相会！"

                today_str = today.strftime("%Y年%m月%d日")
                if is_xmas:
                    text = (f"宝宝，今天是{today_str}第{count}个1314，我们已相恋{duration}，"
                            f"从圣诞夜一直走到今天，未来也要一起闪耀！🎄❤{special}")
                else:
                    text = (f"宝宝，今天是{today_str}第{count}个1314，"
                            f"我们已经相恋了{duration}，爱你❤{special}")
            else:
                # 3. 否则从 Excel 取
                excel_path = params.get("excel_path", "").strip()
                if not excel_path or not os.path.isfile(excel_path):
                    raise FileNotFoundError("未指定或找不到 Excel 文件")

                sheet_id = params.get("sheet", "0")
                col_index = int(params.get("col", 0))
                mode = params.get("mode", "顺序")

                # === 关键：使用 (文件, 表, 列) 作为缓存键 ===
                cache_key = (excel_path, str(sheet_id), col_index)

                # 1. 检查是否已缓存 workbook（避免重复打开）
                wb_cache_key = excel_path
                if wb_cache_key not in self._excel_cache:
                    wb = openpyxl.load_workbook(excel_path, data_only=True)
                    try:
                        ws = wb[int(sheet_id)] if str(sheet_id).isdigit() else wb[sheet_id]
                    except Exception:
                        ws = wb.worksheets[0]
                    rows = list(ws.iter_rows(values_only=True))
                    self._excel_cache[wb_cache_key] = (wb, ws, rows)
                _, _, rows = self._excel_cache[wb_cache_key]

                if not rows:
                    raise ValueError("Excel 表无数据")

                cells = [row[col_index] for row in rows if len(row) > col_index and row[col_index] is not None]
                if not cells:
                    raise ValueError("指定列为空")

                # === 2. 使用 cache_key 管理 cycle ===
                if mode == "顺序":
                    # 初始化类变量（如果还没创建）
                    if not hasattr(self, '_excel_cycle_dict'):
                        self._excel_cycle_dict = {}

                    # 如果该 (文件, 表, 列) 组合没有 cycle，创建一个
                    if cache_key not in self._excel_cycle_dict:
                        self._excel_cycle_dict[cache_key] = itertools.cycle(cells)

                    text = next(self._excel_cycle_dict[cache_key])

                else:  # 随机

                    text = StrongRandomSelector(cells).select_next()
        self._send_text(str(text))

    def _send_text(self, text: str):
        """真正执行文本输入的公共逻辑"""
        self.log_message.emit(self.task_name, f"⌨️ 文本输入: '{text}'")
        try:
            import pyperclip
            pyperclip.copy(text)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            self.log_message.emit(self.task_name, "✅ 文本输入完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 文本输入出错: {str(e)}")
            raise
    def execute_wait(self, params):
        seconds = params.get("seconds", 0)
        if seconds > 0:
            self.log_message.emit(self.task_name, f"⏱️ 等待 {seconds}秒")
            try:
                time.sleep(seconds)
            except Exception as e:
                self.log_message.emit(self.task_name, f"❌ 等待操作出错: {str(e)}")
                raise

    def execute_screenshot(self, params):
        save_path = params.get("save_path", "")
        region = params.get("region", None)

        self.log_message.emit(self.task_name, f"📸 截图保存到: {save_path}")

        try:
            if region:
                x, y, width, height = region
                self.log_message.emit(self.task_name, f"🖼️ 截图区域: x={x}, y={y}, width={width}, height={height}")
                screenshot = pyautogui.screenshot(region=(x, y, width, height))
            else:
                self.log_message.emit(self.task_name, "🖼️ 全屏截图")
                screenshot = pyautogui.screenshot()

            screenshot.save(save_path)
            self.log_message.emit(self.task_name, "✅ 截图保存成功")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 截图操作出错: {str(e)}")
            raise

    def execute_drag(self, params):
        use_image = params.get("use_image", True)
        duration = params.get("duration", 1.0)

        if use_image:
            # 使用图像识别定位起始点
            image_path = params.get("image_path", "")
            offset_x = params.get("offset_x", 0)
            offset_y = params.get("offset_y", 0)
            drag_x = params.get("drag_x", 0)  # 相对拖拽距离
            drag_y = params.get("drag_y", 100)  # 默认向下拖拽100像素
            confidence = params.get("confidence", 0.8)
            timeout = self.timeout

            if not image_path:
                raise ValueError("图像路径不能为空")

            def find_image_center():
                start = time.time()
                while True:
                    pos = pyautogui.locateCenterOnScreen(image_path, confidence=confidence)
                    if pos:
                        return pos
                    if time.time() - start > timeout:
                        return None
                    time.sleep(0.2)

            center = find_image_center()
            if center is None:
                if self.auto_skip_image_timeout:
                    self.log_message.emit(self.task_name, f"⚠️ 在 {timeout}s 内未找到图片: {os.path.basename(image_path)}，自动跳过")
                    return  # ✅ 跳过，不抛异常
                else:
                    raise RuntimeError(f"在 {timeout}s 内未找到图片: {image_path}")

            start_x = center.x + offset_x
            start_y = center.y + offset_y
            end_x = start_x + drag_x
            end_y = start_y + drag_y

        else:
            # 使用直接坐标
            start_x = params.get("start_x", 0)
            start_y = params.get("start_y", 0)
            end_x = params.get("end_x", 0)
            end_y = params.get("end_y", 0)

        self.log_message.emit(self.task_name,
                              f"↔️ 从 ({start_x}, {start_y}) 拖拽到 ({end_x}, {end_y}), 时长: {duration}秒")

        try:
            pyautogui.moveTo(start_x, start_y)
            pyautogui.dragTo(end_x, end_y, duration=duration, button='left')
            self.log_message.emit(self.task_name, "✅ 拖拽操作完成")
        except Exception as e:
            self.log_message.emit(self.task_name, f"❌ 拖拽操作出错: {str(e)}")
            raise


class StepConfigDialog(QDialog):
    def __init__(self, step_data=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("配置步骤")
        self.setMinimumWidth(500)
        self.setWindowIcon(ATIcon.icon())

        layout = QVBoxLayout(self)

        # 步骤类型选择
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("步骤类型:"))
        self.type_combo = QComboBox()
        self.type_combo.addItems(["鼠标点击", "文本输入", "等待", "截图", "拖拽", "鼠标滚轮", "键盘热键", "自动回复"])
        type_layout.addWidget(self.type_combo)
        layout.addLayout(type_layout)

        # 参数配置区域
        self.params_stack = QWidget()
        self.params_layout = QVBoxLayout(self.params_stack)
        self.params_layout.setContentsMargins(0, 10, 0, 0)

        # 创建不同步骤类型的参数面板
        self.mouse_click_panel = self.create_mouse_click_panel()
        self.keyboard_input_panel = self.create_keyboard_input_panel()
        self.wait_panel = self.create_wait_panel()
        self.screenshot_panel = self.create_screenshot_panel()
        self.drag_panel = self.create_drag_panel()
        self.scroll_panel = self.create_mouse_scroll_panel()
        self.hot_keyboard_panel = self.create_hot_keyboard_panel()
        self.ai_reply_panel = self.create_ai_reply_panel()


        # 添加到堆栈
        self.params_layout.addWidget(self.mouse_click_panel)
        self.params_layout.addWidget(self.keyboard_input_panel)
        self.params_layout.addWidget(self.wait_panel)
        self.params_layout.addWidget(self.screenshot_panel)
        self.params_layout.addWidget(self.drag_panel)
        self.params_layout.addWidget(self.scroll_panel)
        self.params_layout.addWidget(self.hot_keyboard_panel)
        self.params_layout.addWidget(self.ai_reply_panel)

        layout.addWidget(self.params_stack)

        # 延时设置
        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("步骤执行后延时(秒):"))
        self.delay_spin = QSpinBox()
        self.delay_spin.setRange(0, 3600)
        self.delay_spin.setValue(0)
        delay_layout.addWidget(self.delay_spin)
        layout.addLayout(delay_layout)

        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        # 连接信号
        self.type_combo.currentIndexChanged.connect(self.update_params_panel)

        # 初始化UI
        self.update_params_panel()

        # 如果有传入步骤数据，填充表单
        if step_data:
            self.load_step_data(step_data)

    # 1. 新增极简滚轮面板
    def create_mouse_scroll_panel(self):
        panel = QWidget()
        layout = QFormLayout(panel)

        # 方向
        self.scroll_direction_combo = QComboBox()
        self.scroll_direction_combo.addItems(["向上滚动", "向下滚动"])
        layout.addRow("滚动方向:", self.scroll_direction_combo)

        # 格数
        self.scroll_clicks_spin = QSpinBox()
        self.scroll_clicks_spin.setRange(1, 100)
        self.scroll_clicks_spin.setValue(3)
        layout.addRow("滚动格数:", self.scroll_clicks_spin)

        return panel

    def create_hot_keyboard_panel(self):
        panel = QWidget()
        layout = QFormLayout(panel)

        # 热键输入框和按钮
        hotkey_layout = QHBoxLayout()
        self.hotkey_input = QLineEdit()
        self.hotkey_input.setPlaceholderText("点击按钮录制热键")
        self.hotkey_input.setReadOnly(True)

        self.record_hotkey_btn = QPushButton("录制热键")
        self.record_hotkey_btn.clicked.connect(self.start_hotkey_recording)

        hotkey_layout.addWidget(self.hotkey_input)
        hotkey_layout.addWidget(self.record_hotkey_btn)

        layout.addRow("热键:", hotkey_layout)

        # 预设热键下拉框（可选）
        self.preset_hotkey_combo = QComboBox()
        self.preset_hotkey_combo.addItems([
            "Ctrl+C", "Ctrl+V", "Ctrl+X", "Ctrl+Z", "Ctrl+A",
            "Ctrl+S", "Ctrl+F", "Alt+Tab", "Ctrl+Alt+Del",
            "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12",
            "Ctrl+F1", "Ctrl+F2", "Ctrl+F3", "Ctrl+F4", "Ctrl+F5",
            "Alt+F4", "Ctrl+Shift+Esc",'Ctrl+Alt+W','Ctrl+Alt+S','Enter','Backspace','Tab'
        ])
        self.hotkey_input.setText("Ctrl+C")
        self._hotkey_value = "Ctrl+C"
        self.preset_hotkey_combo.currentTextChanged.connect(self.on_preset_hotkey_selected)
        layout.addRow("预设热键:", self.preset_hotkey_combo)

        # 额外延迟（ms）
        self.hotkey_delay_spin = QSpinBox()
        self.hotkey_delay_spin.setRange(0, 5000)
        self.hotkey_delay_spin.setValue(100)
        self.hotkey_delay_spin.setSuffix(" ms")
        layout.addRow("执行后延时:", self.hotkey_delay_spin)

        # 存储热键值的隐藏属性
        return panel

    def on_preset_hotkey_selected(self, text):
        """处理预设热键选择事件"""
        if text:
            self.hotkey_input.setText(text)
            self._hotkey_value = text  # 同时更新 _hotkey_value
    def start_hotkey_recording(self):
        """开始录制热键"""
        self.record_hotkey_btn.setText("按下热键...")
        self.record_hotkey_btn.setEnabled(False)
        self.hotkey_input.clear()

        # 启动热键监听
        self.hotkey_listener = keyboard.Listener(
            on_press=self.on_hotkey_press,
            on_release=self.on_hotkey_release
        )
        self.hotkey_listener.start()
        self.current_keys = set()

    def on_hotkey_press(self, key):
        """热键按下事件"""
        # ========== 新增开始 ==========
        # Windows 把 Ctrl+字母 变成控制字符，这里还原成字母
        if (isinstance(key, KeyCode) and key.char and
                '\x00' <= key.char <= '\x1F' and
                Key.ctrl_l in self.current_keys or Key.ctrl_r in self.current_keys):
            # 还原成 Ctrl+字母
            letter = chr(ord(key.char) + 64)  # 0x01 -> 'A'
            self.current_keys.add(KeyCode.from_char(letter.lower()))
            # 不再把原始 \x01 放进集合
            return
        # ========== 新增结束 ==========
        self.current_keys.add(key)
        # 实时显示当前按键组合
        hotkey_str = self.format_hotkey(self.current_keys)
        self.hotkey_input.setText(hotkey_str)

    def on_hotkey_release(self, key):
        """热键释放事件"""
        # 当所有键都释放时，完成录制
        if key in self.current_keys:
            self.current_keys.remove(key)
            print(self.current_keys)
        if not self.current_keys:  # 所有键都已释放
            hotkey_str = self.hotkey_input.text()
            if hotkey_str:
                self._hotkey_value = hotkey_str
                self.record_hotkey_btn.setText("录制热键")
                self.record_hotkey_btn.setEnabled(True)
                if self.hotkey_listener:
                    self.hotkey_listener.stop()
            return False  # 停止监听

    def format_hotkey(self, keys):
        """
        把 pynput 得到的按键列表转成统一字符串，例如：
        [Key.ctrl, Key.alt, KeyCode.from_char('w')]  ->  'CTRL+ALT+W'
        """
        names = []

        for k in keys:
            if isinstance(k, Key):
                # 特殊键：统一大小写并去掉 _l / _r
                name = {
                    Key.ctrl_l: 'CTRL',
                    Key.ctrl_r: 'CTRL',
                    Key.alt_l: 'ALT',
                    Key.alt_r: 'ALT',
                    Key.shift_l: 'SHIFT',
                    Key.shift_r: 'SHIFT',
                    Key.cmd: 'WIN',  # ← 新增这一行
                    Key.cmd_r: 'WIN',  # 右Win 保险起见也写上
                    Key.cmd_l: 'WIN',  # 左Win 保险起见也写上
                }.get(k, k.name.upper())
                names.append(name)

            elif isinstance(k, KeyCode):
                # 普通字符：优先用 char 字段
                char = k.char.upper() if k.char else ''
                if char:
                    names.append(char)
                else:
                    # 功能键、空格、回车等用 vk → 名字映射
                    try:
                        names.append(Key.from_vk(k.vk).name.upper())
                    except ValueError as e:
                        print(f"无法将按键 {k} 转换为名称：{e}")

        # 去重并保持顺序：CTRL/ALT/SHIFT 在前，其余在后
        modifiers = [n for n in names if n in {'CTRL', 'ALT', 'SHIFT','WIN'}]
        others = [n for n in names if n not in {'CTRL', 'ALT', 'SHIFT','WIN'}]

        # 利用 dict.fromkeys 去重并保持首次出现顺序
        ordered = list(dict.fromkeys(modifiers + others))
        return '+'.join(ordered)

    def capture_region(self):
        parent = self.parent()

        parent.hide()
        self.hide()

        self.overlay = RegionCaptureOverlay()
        self.overlay.finished.connect(self.on_region_done)
        self.overlay.show()

    def on_region_done(self, geo: QRect):

        if hasattr(self, 'overlay') and self.overlay is not None:
            self.overlay.close()
            self.overlay.deleteLater()
            self.overlay = None

        parent = self.parent()
        if geo.isNull():
            print("❌ 用户未选择有效区域")
            parent.show()
            self.show()
            return

        pixmap = QApplication.primaryScreen().grabWindow(
            0, geo.x(), geo.y(), geo.width(), geo.height()
        )
        img_dir = os.path.join(os.getcwd(), "img")
        # img_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "img")
        os.makedirs(img_dir, exist_ok=True)
        file_name = datetime.now().strftime("%Y%m%d_%H%M%S") + ".png"
        file_path = os.path.join(img_dir, file_name)

        if pixmap.save(file_path, "PNG"):
            self.image_path_edit.setText(file_path)
            self.drag_image_path_edit.setText(file_path)
            QMessageBox.information(self, "框选截图成功", f"已保存：{file_name}")
            # 直接调用 add_step_to_table
            step_data = self.get_step_data()
            parent.add_step_to_table(step_data)
            # 添加到当前任务配置
            if parent.current_task and parent.current_task in parent.tasks:
                parent.tasks[parent.current_task]["steps"].append(step_data)
        else:
            QMessageBox.warning(self, "失败", "截图保存失败！")

        parent.show()
        self.show()


    def create_mouse_click_panel(self):
        panel = QWidget()
        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        self.dianji_use_image_checkbox = QCheckBox("启用图片")
        self.dianji_use_image_checkbox.setChecked(True)  # 默认
        layout.addWidget(self.dianji_use_image_checkbox, 0, 0)
        layout.addWidget(QLabel("图片路径:"), 0, 1)
        self.image_path_edit = QLineEdit()
        layout.addWidget(self.image_path_edit, 0, 2)
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_image)
        layout.addWidget(browse_btn, 0, 3)

        # >>> 新增：一键录制按钮
        record_btn = QPushButton("框选截图")
        record_btn.clicked.connect(self.capture_region)
        record_btn.setToolTip(
            "请先设置鼠标点击的其他设置\n 如偏移 识别精度 最后再进行框选截图 \n这样才会使得其他设置有效\n（ps:这是个使用bug 待修复）")
        layout.addWidget(record_btn, 0, 4)
        layout.addWidget(record_btn, 0, 4)

        # 坐标输入行
        self.use_coordinate_checkbox = QCheckBox("启用坐标")
        self.use_coordinate_checkbox.setChecked(False)  # 默认

        layout.addWidget(self.use_coordinate_checkbox, 1, 0)
        layout.addWidget(QLabel("X坐标:"),1,1)
        self.x_coordinate_spinbox = QSpinBox()
        self.x_coordinate_spinbox.setRange(0, 100000)
        self.x_coordinate_spinbox.setValue(0)
        layout.addWidget(self.x_coordinate_spinbox,1,2)

        layout.addWidget(QLabel("Y坐标:"),1,3)
        self.y_coordinate_spinbox = QSpinBox()
        self.y_coordinate_spinbox.setRange(0, 100000)
        self.y_coordinate_spinbox.setValue(0)
        layout.addWidget(self.y_coordinate_spinbox,1,4)

        # 创建互斥的按钮组
        self.mode_group = QButtonGroup(self)
        self.mode_group.setExclusive(True)  # 设置为互斥模式
        self.mode_group.addButton(self.dianji_use_image_checkbox)
        self.mode_group.addButton(self.use_coordinate_checkbox)
        self.mode_group.buttonToggled.connect(self.on_mode_changed)

        # 坐标拾取按钮
        self.pick_coordinate_btn = QPushButton("拾取坐标")
        self.pick_coordinate_btn.clicked.connect(self.start_coordinate_picking)
        layout.addWidget(self.pick_coordinate_btn,1,5)

        # 点击类型和读取方向
        layout.addWidget(QLabel("点击类型:"), 2, 0)
        self.click_type_combo = QComboBox()
        self.click_type_combo.addItems(["左键单击", "左键双击", "右键单击", "中键单击"])
        layout.addWidget(self.click_type_combo, 2, 1)

        # 图片读取方向
        layout.addWidget(QLabel("读取方向:"), 2, 2)
        self.scan_direction_combo = QComboBox()
        self.scan_direction_combo.addItems(["默认","从左到右", "从右到左", "从上到下", "从下到上"])
        layout.addWidget(self.scan_direction_combo, 2, 3)

        # 偏移量
        layout.addWidget(QLabel("X偏移:"), 3, 0)
        self.offset_x_spin = QSpinBox()
        self.offset_x_spin.setRange(-1000, 1000)
        layout.addWidget(self.offset_x_spin, 3, 1)

        layout.addWidget(QLabel("Y偏移:"), 3, 2)
        self.offset_y_spin = QSpinBox()
        self.offset_y_spin.setRange(-1000, 1000)
        layout.addWidget(self.offset_y_spin, 3, 3)

        # 识别设置
        layout.addWidget(QLabel("识别精度(0-1):"), 4, 0)
        self.confidence_spin = QDoubleSpinBox()
        self.confidence_spin.setRange(0.5, 1.0)
        self.confidence_spin.setValue(0.8)
        self.confidence_spin.setSingleStep(0.05)
        layout.addWidget(self.confidence_spin, 4, 1)

        layout.addWidget(QLabel("超时时间(秒):"), 4, 2)
        self.timeout_spin = QDoubleSpinBox()
        self.timeout_spin.setRange(0.1, 60)
        self.timeout_spin.setSingleStep(0.1)
        self.timeout_spin.setValue(1.0)
        self.timeout_spin.setDecimals(1)
        layout.addWidget(self.timeout_spin, 4, 3)

        return panel

    def on_mode_changed(self, button, checked):
        """模式切换处理"""
        if checked:
            self.update_controls_state()

    def update_controls_state(self):
        """更新控件启用状态"""
        image_enabled = self.dianji_use_image_checkbox.isChecked()
        coordinate_enabled = self.use_coordinate_checkbox.isChecked()

        # 更新图片相关控件状态
        self.image_path_edit.setEnabled(image_enabled)

        # 更新坐标相关控件状态
        self.x_coordinate_spinbox.setEnabled(coordinate_enabled)
        self.y_coordinate_spinbox.setEnabled(coordinate_enabled)

    def start_coordinate_picking(self):
        """
        开始坐标拾取
        """

        self.coord_picker = CoordinatePickerOverlay(self)
        self.coord_picker.coordinate_selected.connect(self.on_coordinate_selected)
        self.coord_picker.finished.connect(self.on_coordinate_picking_finished)
        # 创建并显示坐标拾取覆盖层
        parent = self.parent()
        parent.showMinimized()
        self.coord_picker.show()
        self.coord_picker.raise_()
        self.coord_picker.activateWindow()

    def on_coordinate_selected(self, coordinate):
        """
        坐标选择完成的回调
        """
        x, y = coordinate
        self.x_coordinate_spinbox.setValue(x)
        self.y_coordinate_spinbox.setValue(y)

        # 如果当前是使用坐标模式，更新预览
        if not self.dianji_use_image_checkbox.isChecked():
            self.update_mouse_click_preview()
        parent = self.parent()
        parent.showMinimized()

    def on_coordinate_picking_finished(self):
        """
        坐标拾取完成后的处理
        """
        # 清理引用
        self.coord_picker.deleteLater()
        self.coord_picker = None

        # 显示主窗口
        parent = self.parent()
        parent.showNormal()
        self.raise_()
        self.activateWindow()

    def update_mouse_click_preview(self):
        """
        更新鼠标点击预览
        """
        # 这里可以添加预览逻辑，如果需要的话
        pass

    def create_ai_reply_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # AI 提供商选择
        provider_layout = QHBoxLayout()
        provider_layout.addWidget(QLabel("AI 提供商:"))
        self.ai_provider_combo = QComboBox()
        self.ai_provider_combo.addItems(["kimi", "doubao"])
        provider_layout.addWidget(self.ai_provider_combo)
        layout.addLayout(provider_layout)

        # 预设角色选择
        role_layout = QHBoxLayout()
        role_layout.addWidget(QLabel("预设角色:"))
        self.role_combo = QComboBox()
        self.role_combo.addItems([
            "自定义",
            "贴心朋友 ❤️",
            "幽默损友 😂",
            "专业助手 🧠",
            "温柔恋人 💕",
            "严厉老师 🎓"
        ])
        role_layout.addWidget(self.role_combo)
        layout.addLayout(role_layout)

        # 系统提示词
        system_prompt_layout = QVBoxLayout()
        system_prompt_layout.addWidget(QLabel("系统提示词:"))
        self.ai_system_prompt_edit = QTextEdit()  # 改名以避免冲突
        self.ai_system_prompt_edit.setMaximumHeight(150)
        self.ai_system_prompt_edit.setPlaceholderText(
            "请输入系统提示词，例如：\n你是我的朋友，微信语音里很随和。用一句口语化的话回应我...")
        system_prompt_layout.addWidget(self.ai_system_prompt_edit)
        layout.addLayout(system_prompt_layout)

        # 预设角色提示词
        self.role_prompts = {
            "自定义": "",
            "贴心朋友 ❤️": (
                "你是关系不错的朋友，语气温暖但不过度热情。先简短共情，再给一个贴心小建议或小问题，"
                "全程口语化、一句话为主，不提自己是AI，不使用列表。"
            ),
            "幽默损友 😂": (
                "你是损友风格的朋友，聊天轻松搞笑但不冒犯。用一句俏皮话回应，"
                "可以小小自嘲或反转，像真人发语音那样自然，不暴露身份。"
            ),
            "专业助手 🧠": (
                "你是一位专业的助手，回答问题时要准确、清晰、有条理。"
                "使用正式但友好的语气，提供有用的信息和建议，避免冗长。"
            ),
            "温柔恋人 💕": (
                "你是我温柔的恋人，说话充满关爱和体贴。"
                "用甜蜜温馨的话语回应，适当使用表情符号表达情感，让对话充满爱意。"
            ),
            "严厉老师 🎓": (
                "你是一位严格的老师，语气认真且专业。"
                "回答问题时要指出关键要点，适当纠正错误，给出建设性的建议和指导。"
            )
        }

        # 连接角色选择信号
        self.role_combo.currentTextChanged.connect(self.on_ai_role_changed)

        # 初始化为贴心朋友角色
        # self.role_combo.setCurrentText("贴心朋友 ❤️")
        # self.on_ai_role_changed("贴心朋友 ❤️")

        # 历史记录选项
        history_layout = QHBoxLayout()
        self.use_history_checkbox = QCheckBox("使用对话历史")
        self.use_history_checkbox.setChecked(True)
        history_layout.addWidget(self.use_history_checkbox)

        # 添加剪切板输入选项
        self.use_clipboard_checkbox = QCheckBox("使用剪切板作为消息输入")
        self.use_clipboard_checkbox.setChecked(False)  # 默认不启用
        history_layout.addWidget(self.use_clipboard_checkbox)

        layout.addLayout(history_layout)

        # 流式输出选项
        stream_layout = QHBoxLayout()
        self.stream_checkbox = QCheckBox("流式输出")
        self.stream_checkbox.setChecked(False)
        stream_layout.addWidget(self.stream_checkbox)
        layout.addLayout(stream_layout)

        return panel

    def on_ai_role_changed(self, role_text):
        """处理AI角色选择变化"""
        # 检查控件是否仍然存在
        if not hasattr(self, 'ai_system_prompt_edit'):
            return
        try:
            if role_text in self.role_prompts:
                prompt = self.role_prompts[role_text]
                self.ai_system_prompt_edit.setPlainText(prompt)
                # 如果是自定义角色，允许用户编辑
                # self.ai_system_prompt_edit.setReadOnly(role_text != "自定义")
        except RuntimeError as e:
            # 控件已被删除，忽略错误
            print(e)
            pass
    def generate_love_text(self):
        from datetime import datetime, date, time
        love_dt = self.love_datetime_edit.dateTime().toPython()  # 用户选的时刻
        today = date.today()
        today_1314 = datetime.combine(today, time(13, 14))  # 今天 13:14

        # 相恋时长（精确到秒）
        delta = today_1314 - love_dt
        days = delta.days
        sec = delta.seconds
        hours, rem = divmod(sec, 3600)
        minutes, secs = divmod(rem, 60)
        duration = f"{days}天{hours}时{minutes}分{secs}秒"

        # 今年第几个 13:14
        year_start_1314 = datetime(today.year, 1, 1, 13, 14)
        count = (today_1314 - year_start_1314).days + 1

        # 特殊节日
        year = today.year
        is_xmas = (love_dt.month, love_dt.day) == (12, 25)
        special = None
        if is_xmas:
            special = "我们的爱情从圣诞夜点亮，愿它像圣诞树一样永远闪耀！"
        elif today == date(year, 2, 14):
            special = "情人节快乐！"
        elif today == self.chinese_qixi(year):
            special = "七夕快乐，鹊桥相会！"
        elif today == date(year, 12, 25):
            special = "圣诞快乐，Merry Christmas！"

        today_str = f"{today.year}年{today.month}月{today.day}日"
        if is_xmas:
            text = (f"宝宝，今天是{today_str}第{count}个1314，"
                    f"我们已相恋{duration}，"
                    f"从圣诞夜一直走到今天，未来也要一起闪耀！🎄❤")
        else:
            text = (f"宝宝，今天是{today_str}第{count}个1314，"
                    f"我们已经相恋了{duration}，爱你❤ ")
            if special:
                text += f"\n{special}"

        self.text_edit.setPlainText(text)
    def create_keyboard_input_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # 1. 原始文本输入（多行）
        layout.addWidget(QLabel("输入文本（留空则用 Excel 或纪念日）:"))

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText("在此输入固定文本...\n留空则自动从 Excel 或纪念日生成内容")
        self.text_edit.setMaximumHeight(80)
        self.text_edit.setLineWrapMode(QPlainTextEdit.WidgetWidth)

        layout.addWidget(self.text_edit)


        # -------- 新增纪念日区域 --------
        love_group = QWidget()
        h_layout = QHBoxLayout(love_group)  # 横向布局

        # 1. 启用复选框
        self.use_love_checkbox = QCheckBox("启用纪念日")
        self.use_love_checkbox.setChecked(False)  # 默认不启用
        h_layout.addWidget(self.use_love_checkbox)

        # 2. 标签
        h_layout.addWidget(QLabel("时间:"))

        # 3. 时间选择器
        self.love_datetime_edit = QDateTimeEdit()
        self.love_datetime_edit.setCalendarPopup(True)
        self.love_datetime_edit.setDisplayFormat("yyyy-MM-dd hh:mm:ss")
        self.love_datetime_edit.setDateTime(QDateTime(QDate(2022, 12, 25), QTime(7, 0, 0)))
        # 可选：默认禁用，直到 checkbox 勾选
        self.love_datetime_edit.setEnabled(False)
        self.use_love_checkbox.toggled.connect(self.love_datetime_edit.setEnabled)

        h_layout.addWidget(self.love_datetime_edit)
        # 4. 生成按钮
        gen_btn = QPushButton("生成文案")
        gen_btn.setEnabled(False)
        self.use_love_checkbox.toggled.connect(gen_btn.setEnabled)  # 勾选/取消自动启用/禁用
        gen_btn.clicked.connect(self.generate_love_text)
        h_layout.addWidget(gen_btn)

        # 可选：设置拉伸，防止挤压
        h_layout.addStretch()

        # 将 group 添加到主 layout
        layout.addWidget(love_group)


        # 3. Excel 区域
        excel_group = QWidget()
        g = QVBoxLayout(excel_group)

        # 文件选择
        h1 = QHBoxLayout()
        h1.addWidget(QLabel("Excel 文件:"))
        self.excel_path_edit = QLineEdit()
        btn = QPushButton("浏览")
        btn.clicked.connect(lambda: self.excel_path_edit.setText(
            QFileDialog.getOpenFileName(filter="*.xlsx")[0]))
        h1.addWidget(self.excel_path_edit)
        h1.addWidget(btn)
        g.addLayout(h1)

        # 工作表
        h2 = QHBoxLayout()
        h2.addWidget(QLabel("工作表(名称或序号):"))
        self.sheet_edit = QLineEdit("0")
        h2.addWidget(self.sheet_edit)
        g.addLayout(h2)

        # 列
        h3 = QHBoxLayout()
        h3.addWidget(QLabel("列(首列=0):"))
        self.col_spin = QSpinBox()
        self.col_spin.setValue(0)
        h3.addWidget(self.col_spin)
        g.addLayout(h3)

        # 读取模式
        h4 = QHBoxLayout()
        h4.addWidget(QLabel("读取模式:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["顺序", "随机"])
        h4.addWidget(self.mode_combo)
        g.addLayout(h4)

        layout.addWidget(excel_group)
        return panel


    def create_wait_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)

        layout.addWidget(QLabel("等待时间(秒):"))
        self.wait_spin = QSpinBox()
        self.wait_spin.setRange(1, 3600)
        self.wait_spin.setValue(5)
        layout.addWidget(self.wait_spin)

        return panel

    def create_screenshot_panel(self):
        panel = QWidget()
        layout = QGridLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 保存路径
        layout.addWidget(QLabel("保存路径:"), 0, 0)
        self.screenshot_path_edit = QLineEdit()
        layout.addWidget(self.screenshot_path_edit, 0, 1)
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(self.browse_save_path)
        layout.addWidget(browse_btn, 0, 2)

        # 截图区域
        layout.addWidget(QLabel("截图区域(可选):"), 1, 0)

        layout.addWidget(QLabel("X:"), 2, 0)
        self.screenshot_x_spin = QSpinBox()
        self.screenshot_x_spin.setRange(0, 10000)
        layout.addWidget(self.screenshot_x_spin, 2, 1)

        layout.addWidget(QLabel("Y:"), 2, 2)
        self.screenshot_y_spin = QSpinBox()
        self.screenshot_y_spin.setRange(0, 10000)
        layout.addWidget(self.screenshot_y_spin, 2, 3)

        layout.addWidget(QLabel("宽度:"), 3, 0)
        self.screenshot_width_spin = QSpinBox()
        self.screenshot_width_spin.setRange(1, 10000)
        self.screenshot_width_spin.setValue(800)
        layout.addWidget(self.screenshot_width_spin, 3, 1)

        layout.addWidget(QLabel("高度:"), 3, 2)
        self.screenshot_height_spin = QSpinBox()
        self.screenshot_height_spin.setRange(1, 10000)
        self.screenshot_height_spin.setValue(600)
        layout.addWidget(self.screenshot_height_spin, 3, 3)

        return panel

    # 在 StepConfigDialog 类中添加新的拖拽面板
    def create_drag_panel(self):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)

        # 添加图像识别选项
        self.use_image_checkbox = QCheckBox("使用图像识别定位起始点")
        self.use_image_checkbox.setChecked(True)
        layout.addWidget(self.use_image_checkbox)

        # 图像路径设置
        image_layout = QHBoxLayout()
        image_layout.addWidget(QLabel("起始点图像:"))
        self.drag_image_path_edit = QLineEdit()
        image_browse_btn = QPushButton("浏览...")
        image_browse_btn.clicked.connect(self.browse_drag_image)

        # >>> 新增：一键录制按钮
        record_btn = QPushButton("框选截图")
        record_btn.clicked.connect(self.capture_region)

        image_layout.addWidget(self.drag_image_path_edit)
        image_layout.addWidget(image_browse_btn)
        image_layout.addWidget(record_btn)
        layout.addLayout(image_layout)

        # 偏移量设置
        offset_layout = QHBoxLayout()
        offset_layout.addWidget(QLabel("图像识别偏移:"))
        offset_layout.addWidget(QLabel("X:"))
        self.drag_offset_x_spin = QSpinBox()
        self.drag_offset_x_spin.setRange(-1000, 1000)
        offset_layout.addWidget(self.drag_offset_x_spin)

        offset_layout.addWidget(QLabel("Y:"))
        self.drag_offset_y_spin = QSpinBox()
        self.drag_offset_y_spin.setRange(-1000, 1000)
        offset_layout.addWidget(self.drag_offset_y_spin)

        offset_layout.addWidget(QLabel("读取方向:"))
        self.drag_scan_direction_combo = QComboBox()
        self.drag_scan_direction_combo.addItems(["默认","从左到右", "从右到左", "从上到下", "从下到上"])
        offset_layout.addWidget(self.drag_scan_direction_combo)
        offset_layout.addStretch()
        layout.addLayout(offset_layout)

        # 拖拽距离（相对拖拽）
        distance_layout = QHBoxLayout()
        distance_layout.addWidget(QLabel("横向距离:"))
        self.drag_distance_x_spin = QSpinBox()
        self.drag_distance_x_spin.setRange(-1000, 1000)
        self.drag_distance_x_spin.setValue(0)
        distance_layout.addWidget(self.drag_distance_x_spin)

        distance_layout.addWidget(QLabel("纵向距离:"))
        self.drag_distance_y_spin = QSpinBox()
        self.drag_distance_y_spin.setRange(-1000, 1000)
        self.drag_distance_y_spin.setValue(100)  # 默认向下拖拽100像素
        distance_layout.addWidget(self.drag_distance_y_spin)

        # 添加快捷按钮
        up_btn = QPushButton("↑上拉")
        up_btn.setFixedSize(60, 25)
        up_btn.clicked.connect(lambda: self.set_drag_distance(0, -100))
        distance_layout.addWidget(up_btn)

        down_btn = QPushButton("↓下拉")
        down_btn.setFixedSize(60, 25)
        down_btn.clicked.connect(lambda: self.set_drag_distance(0, 100))
        distance_layout.addWidget(down_btn)

        left_btn = QPushButton("←左拉")
        left_btn.setFixedSize(60, 25)
        left_btn.clicked.connect(lambda: self.set_drag_distance(-100, 0))
        distance_layout.addWidget(left_btn)

        right_btn = QPushButton("→右拉")
        right_btn.setFixedSize(60, 25)
        right_btn.clicked.connect(lambda: self.set_drag_distance(100, 0))
        distance_layout.addWidget(right_btn)

        distance_layout.addStretch()
        layout.addLayout(distance_layout)

        # 识别设置
        recognition_layout = QHBoxLayout()
        recognition_layout.addWidget(QLabel("识别精度(0-1):"))
        self.drag_confidence_spin = QDoubleSpinBox()
        self.drag_confidence_spin.setRange(0.5, 1.0)
        self.drag_confidence_spin.setValue(0.8)
        self.drag_confidence_spin.setSingleStep(0.05)
        recognition_layout.addWidget(self.drag_confidence_spin)

        recognition_layout.addWidget(QLabel("超时时间(秒):"))
        self.drag_timeout_spin = QDoubleSpinBox()
        self.drag_timeout_spin.setRange(0.1, 60)
        self.drag_timeout_spin.setSingleStep(0.1)
        self.drag_timeout_spin.setValue(10.0)
        self.drag_timeout_spin.setDecimals(1)
        recognition_layout.addWidget(self.drag_timeout_spin)
        layout.addLayout(recognition_layout)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # 直接坐标设置（可选）
        direct_coords_group = QGroupBox("或直接设置坐标")
        self.direct_coords_group = direct_coords_group
        direct_layout = QHBoxLayout()

        # 起点坐标
        start_layout = QHBoxLayout()
        start_layout.addWidget(QLabel("起点:"))
        start_layout.addWidget(QLabel("X:"))
        self.drag_start_x_spin = QSpinBox()
        self.drag_start_x_spin.setRange(0, 10000)
        start_layout.addWidget(self.drag_start_x_spin)

        start_layout.addWidget(QLabel("Y:"))
        self.drag_start_y_spin = QSpinBox()
        self.drag_start_y_spin.setRange(0, 10000)
        start_layout.addWidget(self.drag_start_y_spin)

        # 终点坐标
        end_layout = QHBoxLayout()
        end_layout.addWidget(QLabel("终点:"))
        end_layout.addWidget(QLabel("X:"))
        self.drag_end_x_spin = QSpinBox()
        self.drag_end_x_spin.setRange(0, 10000)
        end_layout.addWidget(self.drag_end_x_spin)

        end_layout.addWidget(QLabel("Y:"))
        self.drag_end_y_spin = QSpinBox()
        self.drag_end_y_spin.setRange(0, 10000)
        end_layout.addWidget(self.drag_end_y_spin)

        direct_layout.addLayout(start_layout)
        direct_layout.addLayout(end_layout)
        direct_coords_group.setLayout(direct_layout)
        layout.addWidget(direct_coords_group)


        # 拖拽时间
        time_layout = QHBoxLayout()
        time_layout.addWidget(QLabel("拖拽时间(秒):"))
        self.drag_duration_spin = QDoubleSpinBox()
        self.drag_duration_spin.setRange(0.1, 10.0)
        self.drag_duration_spin.setValue(1.0)
        self.drag_duration_spin.setSingleStep(0.1)
        time_layout.addWidget(self.drag_duration_spin)
        layout.addLayout(time_layout)

        # 连接信号
        self.use_image_checkbox.toggled.connect(self.toggle_drag_mode)
        self.toggle_drag_mode(True)

        return panel

    def set_drag_distance(self, x_distance, y_distance):
        """设置拖拽距离的快捷方法"""
        self.drag_distance_x_spin.setValue(x_distance)
        self.drag_distance_y_spin.setValue(y_distance)

    def toggle_drag_mode(self, use_image):
        """切换拖拽模式"""
        self.direct_coords_group.setDisabled(use_image)

    def browse_drag_image(self):
        """浏览拖拽起始点图像"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择起始点图像", "", "图片文件 (*.png *.jpg *.bmp)"
        )
        if file_path:
            self.drag_image_path_edit.setText(file_path)

    def browse_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择图片", "", "图片文件 (*.png *.jpg *.bmp)"
        )
        if file_path:
            self.image_path_edit.setText(file_path)

    def browse_save_path(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存截图", "", "PNG图片 (*.png)"
        )
        if file_path:
            if not file_path.lower().endswith('.png'):
                file_path += '.png'
            self.screenshot_path_edit.setText(file_path)

    def update_params_panel(self):
        # 隐藏所有面板
        for i in range(self.params_layout.count()):
            widget = self.params_layout.itemAt(i).widget()
            if widget:
                widget.hide()

        # 显示当前选中的面板
        step_type = self.type_combo.currentText()
        if step_type == "鼠标点击":
            self.mouse_click_panel.show()
        elif step_type == "文本输入":
            self.keyboard_input_panel.show()
        elif step_type == "等待":
            self.wait_panel.show()
        elif step_type == "截图":
            self.screenshot_panel.show()
        elif step_type == "拖拽":
            self.drag_panel.show()
        elif step_type == "鼠标滚轮":
            self.scroll_panel.show()
        elif step_type == "键盘热键":
            self.hot_keyboard_panel.show()
        elif step_type == "自动回复":  # 新增
            self.ai_reply_panel.setVisible(True)
            # 隐藏其他面板

    def load_step_data(self, step_data):
        step_type = step_data.get("type", "")
        self.type_combo.setCurrentText(step_type)

        # 设置延时
        self.delay_spin.setValue(step_data.get("delay", 0))

        # 设置参数
        params = step_data.get("params", {})
        if step_type == "鼠标点击":
            self.dianji_use_image_checkbox.setChecked(params.get("use_image", True))
            self.use_coordinate_checkbox.setChecked(params.get("use_coordinates", False))

            self.image_path_edit.setText(params.get("image_path", ""))
            self.click_type_combo.setCurrentText(params.get("click_type", "左键单击"))
            self.scan_direction_combo.setCurrentText(params.get("scan_direction", "默认"))
            self.offset_x_spin.setValue(params.get("offset_x", 0))
            self.offset_y_spin.setValue(params.get("offset_y", 0))
            self.confidence_spin.setValue(params.get("confidence", 0.8))
            self.timeout_spin.setValue(params.get("timeout", 10))
            self.x_coordinate_spinbox.setValue(params.get("x_coordinate", 0))
            self.y_coordinate_spinbox.setValue(params.get("y_coordinate", 0))

        elif step_type == "文本输入":
            self.text_edit.setPlainText(params.get("text", ""))
            self.excel_path_edit.setText(params.get("excel_path", ""))
            self.sheet_edit.setText(str(params.get("sheet", "0")))
            self.col_spin.setValue(int(params.get("col", 0)))
            # 确保下拉框里能找到对应文本
            mode = params.get("mode", "顺序")
            if mode in ["顺序", "随机"]:
                self.mode_combo.setCurrentText(mode)
            else:
                self.mode_combo.setCurrentIndex(0)  # 默认顺序
        elif step_type == "等待":
            self.wait_spin.setValue(params.get("seconds", 5))
        elif step_type == "截图":
            self.screenshot_path_edit.setText(params.get("save_path", ""))
            region = params.get("region", [0, 0, 0, 0])
            if len(region) == 4:
                self.screenshot_x_spin.setValue(region[0])
                self.screenshot_y_spin.setValue(region[1])
                self.screenshot_width_spin.setValue(region[2])
                self.screenshot_height_spin.setValue(region[3])
        elif step_type == "拖拽":
            use_image = params.get("use_image", True)
            self.use_image_checkbox.setChecked(use_image)
            if use_image:
                self.drag_image_path_edit.setText(params.get("image_path", ""))
                self.drag_offset_x_spin.setValue(params.get("offset_x", 0))
                self.drag_offset_y_spin.setValue(params.get("offset_y", 0))
                self.scan_direction_combo.setCurrentText(params.get("scan_direction", "默认"))
                self.drag_distance_x_spin.setValue(params.get("drag_x", 0))
                self.drag_distance_y_spin.setValue(params.get("drag_y", 100))
                self.drag_confidence_spin.setValue(params.get("confidence", 0.8))
                self.drag_timeout_spin.setValue(params.get("timeout", 10.0))
            else:
                self.drag_start_x_spin.setValue(params.get("start_x", 0))
                self.drag_start_y_spin.setValue(params.get("start_y", 0))
                self.drag_end_x_spin.setValue(params.get("end_x", 0))
                self.drag_end_y_spin.setValue(params.get("end_y", 0))
            self.drag_duration_spin.setValue(params.get("duration", 1.0))
        elif step_type == "鼠标滚轮":
            self.scroll_direction_combo.setCurrentText(params.get("direction", "向下滚动"))
            self.scroll_clicks_spin.setValue(params.get("clicks", 3))
        elif step_type == "键盘热键":
            hotkey = params.get("hotkey", "ctrl+c").upper()
            self.hotkey_input.setText(hotkey)
            self._hotkey_value = hotkey
            self.hotkey_delay_spin.setValue(params.get("delay_ms", 100))
        elif step_type == "自动回复":
            self.ai_provider_combo.setCurrentText(params.get("provider", "kimi"))
            self.ai_system_prompt_edit.setPlainText(params.get("system_prompt", ""))
            self.use_history_checkbox.setChecked(params.get("use_history", True))
            self.use_clipboard_checkbox.setChecked(params.get("use_clipboard", True))
            self.stream_checkbox.setChecked(params.get("stream", False))
            # 设置角色下拉框，如果提示词匹配预设角色
            current_prompt = params.get("system_prompt", "")
            for role, prompt in self.role_prompts.items():
                if current_prompt == prompt:
                    self.role_combo.setCurrentText(role)
                    break
            else:
                self.role_combo.setCurrentText("自定义")
    def get_step_data(self):
        step_type = self.type_combo.currentText()
        params = {}

        if step_type == "鼠标点击":
            params = {
                "use_image": self.dianji_use_image_checkbox.isChecked(),
                "image_path": self.image_path_edit.text(),
                "click_type": self.click_type_combo.currentText(),
                "scan_direction": self.scan_direction_combo.currentText(),
                "offset_x": self.offset_x_spin.value(),
                "offset_y": self.offset_y_spin.value(),
                "confidence": self.confidence_spin.value(),
                "timeout": self.timeout_spin.value(),
                "use_coordinates": self.use_coordinate_checkbox.isChecked(),
                "x_coordinate": self.x_coordinate_spinbox.value(),
                "y_coordinate": self.y_coordinate_spinbox.value()
            }
        elif step_type == "文本输入":
            use_love = self.use_love_checkbox.isChecked()
            love_date_str = ""
            if use_love:
                love_date_str = self.love_datetime_edit.dateTime().toPython().isoformat()
            params = {
                "text": self.text_edit.toPlainText().strip(),
                "excel_path": self.excel_path_edit.text().strip(),
                "sheet": self.sheet_edit.text().strip(),
                "col": self.col_spin.value(),
                "mode": self.mode_combo.currentText(),
                "love_date": love_date_str  # 只有启用时才传值
            }
        elif step_type == "等待":
            params = {
                "seconds": self.wait_spin.value()
            }
        elif step_type == "截图":
            params = {
                "save_path": self.screenshot_path_edit.text(),
                "region": [
                    self.screenshot_x_spin.value(),
                    self.screenshot_y_spin.value(),
                    self.screenshot_width_spin.value(),
                    self.screenshot_height_spin.value()
                ]
            }
        elif step_type == "拖拽":
            use_image = self.use_image_checkbox.isChecked()
            params = {
                "use_image": use_image,
                "duration": self.drag_duration_spin.value()
            }
            if use_image:
                params.update({
                    "image_path": self.drag_image_path_edit.text(),
                    "offset_x": self.drag_offset_x_spin.value(),
                    "offset_y": self.drag_offset_y_spin.value(),
                "scan_direction": self.scan_direction_combo.currentText(),
                    "drag_x": self.drag_distance_x_spin.value(),
                    "drag_y": self.drag_distance_y_spin.value(),
                    "confidence": self.drag_confidence_spin.value(),
                    "timeout": self.drag_timeout_spin.value()
                })
            else:
                params.update({
                    "start_x": self.drag_start_x_spin.value(),
                    "start_y": self.drag_start_y_spin.value(),
                    "end_x": self.drag_end_x_spin.value(),
                    "end_y": self.drag_end_y_spin.value()
                })
        elif step_type == "鼠标滚轮":
            params = {
                "direction": self.scroll_direction_combo.currentText(),
                "clicks": self.scroll_clicks_spin.value()
            }
        elif step_type == "键盘热键":
            params = {
                "hotkey": self._hotkey_value,  # 使用存储的热键值
                "delay_ms": self.hotkey_delay_spin.value()
            }
        elif step_type == "自动回复":
            params = {
                "provider": self.ai_provider_combo.currentText(),
                "system_prompt": self.ai_system_prompt_edit.toPlainText(),
                "use_history": self.use_history_checkbox.isChecked(),
                "use_clipboard": self.use_clipboard_checkbox.isChecked(),
                "stream": self.stream_checkbox.isChecked()
            }
        params["step_time"] = datetime.now().strftime("%H:%M:%S")
        print(f"步骤数据: {params}")
        return {
            "type": step_type,
            "params": params,
            "delay": self.delay_spin.value()
        }

    def chinese_qixi(self,year: int) -> date:
        """
        计算指定年份的七夕节（农历七月初七）的公历日期
        使用近似算法，误差在±1天内

        Args:
            year: 要计算的年份

        Returns:
            该年份七夕节的公历日期
        """
        # 扩展的年份对照表（2000-2030年）
        table = {
            2000: date(2000, 8, 6), 2001: date(2001, 8, 25), 2002: date(2002, 8, 15),
            2003: date(2003, 8, 4), 2004: date(2004, 8, 22), 2005: date(2005, 8, 11),
            2006: date(2006, 7, 31), 2007: date(2007, 8, 19), 2008: date(2008, 8, 7),
            2009: date(2009, 8, 26), 2010: date(2010, 8, 16), 2011: date(2011, 8, 6),
            2012: date(2012, 8, 23), 2013: date(2013, 8, 13), 2014: date(2014, 8, 2),
            2015: date(2015, 8, 20), 2016: date(2016, 8, 9), 2017: date(2017, 8, 28),
            2018: date(2018, 8, 17), 2019: date(2019, 8, 7), 2020: date(2020, 8, 25),
            2021: date(2021, 8, 14), 2022: date(2022, 8, 4), 2023: date(2023, 8, 22),
            2024: date(2024, 8, 10), 2025: date(2025, 8, 1), 2026: date(2026, 8, 19),
            2027: date(2027, 8, 8), 2028: date(2028, 7, 28), 2029: date(2029, 8, 16),
            2030: date(2030, 8, 5)
        }

        # 如果在已知年份范围内，直接返回表中日期
        if year in table:
            return table[year]

        # 对于表外的年份，使用近似算法计算
        # 基础年份选择2023年，七夕日期为8月22日
        base_year = 2023
        base_date = date(base_year, 8, 22)

        # 计算与基础年份的差异（考虑农历年的平均长度）
        year_diff = year - base_year
        # 农历年平均长度约为29.53天×12个月 = 354.36天
        days_diff = round(year_diff * 354.36 - year_diff * 365.25)

        # 计算预估日期
        estimated_date = base_date + timedelta(days=days_diff)

        # 调整到8月附近（七夕通常在7月底到8月底之间）
        if estimated_date.month < 7:
            estimated_date += timedelta(days=30)
        elif estimated_date.month > 9:
            estimated_date -= timedelta(days=30)

        return estimated_date


class TaskItemWidget(QWidget):
    def __init__(self, name, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.task_name = name
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSpacing(8)

        # 任务名称标签 - 设置为透明
        self.name_label = QLabel(name)
        self.name_label.setFont(QFont("Arial", 10, QFont.Medium))
        self.name_label.setMinimumWidth(150)
        self.name_label.setStyleSheet("background: transparent;")  # 设置透明背景

        # 状态标签
        self.status_label = QLabel("已停止")
        self.status_label.setFont(QFont("Arial", 9))
        self.status_label.setStyleSheet("background: transparent;")  # 设置透明背景

        # 操作按钮 - 添加emoji
        self.start_btn = QPushButton("▶️")
        self.start_btn.setToolTip("开始任务")
        self.start_btn.setFixedSize(28, 28)

        self.stop_btn = QPushButton("⏹️")
        self.stop_btn.setToolTip("停止任务")
        self.stop_btn.setFixedSize(28, 28)
        # self.stop_btn.setEnabled(False)

        self.delete_btn = QPushButton("🗑️")
        self.delete_btn.setToolTip("删除任务")
        self.delete_btn.setFixedSize(28, 28)

        # 添加到布局
        layout.addWidget(self.name_label)
        layout.addWidget(self.status_label)
        layout.addStretch()
        layout.addWidget(self.start_btn)
        layout.addWidget(self.stop_btn)
        layout.addWidget(self.delete_btn)

        # 连接信号
        self.start_btn.clicked.connect(self.start_task)
        self.stop_btn.clicked.connect(self.stop_task)
        self.delete_btn.clicked.connect(lambda: self.parent.delete_task(name))

    def start_task(self):
        self.status_label.setText("运行中")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        # 更新主界面状态
        if self.parent:
            self.parent.task_status.setText("运行中")
            self.parent.start_current_task()

    def stop_task(self):
        self.status_label.setText("已停止")
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        # 更新主界面状态
        if self.parent:
            # 检查是否是当前任务且正在定时
            if (self.parent.current_task == self.task_name and
                    self.task_name in self.parent.scheduled_timers):
                self.parent.stop_current_task()
            elif self.parent.current_task == self.task_name:
                self.parent.task_status.setText("已停止")
                self.parent.stop_current_task()


class AutomationUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("自动化任务管理器")

        self.setGeometry(100, 100, 1100, 550)  # 减少高度

        # 应用设置
        self.settings = QSettings("MyCompany", "AutomationManager")
        self.load_settings()

        # 存储任务配置
        self.tasks = {}
        self.current_task = None
        self.task_runner = None
        self.task_thread = None
        self.scheduled_timers = {}  # 存储定时任务的计时器
        # 热键监听器
        self.hotkey_listener = None

        self.setup_hotkey_listener()

        # 创建主布局
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 使用分割器实现可调整宽度的任务列表
        self.splitter = QSplitter(Qt.Horizontal)

        # 左侧任务列表区域
        left_panel = QFrame()
        left_panel.setFrameShape(QFrame.StyledPanel)
        left_panel.setMinimumWidth(280)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(10, 10, 10, 10)

        # 任务列表标题和新建按钮 - 添加emoji
        title_layout = QHBoxLayout()
        title_label = QLabel("📋 任务列表")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_layout.addWidget(title_label)
        title_layout.addStretch()

        self.new_task_btn = QPushButton("➕ 新建任务")
        self.new_task_btn.setFixedSize(100, 32)
        title_layout.addWidget(self.new_task_btn)

        left_layout.addLayout(title_layout)

        # 分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        left_layout.addWidget(separator)

        # 任务列表
        self.task_list = QListWidget()
        self.task_list.setMinimumHeight(200)
        # 优化hover样式
        self.task_list.setStyleSheet("""
            QListWidget::item:hover {
                background-color: #e0e0e0;
            }
        """)

        # 设置右键菜单
        self.task_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.task_list.customContextMenuRequested.connect(self.show_context_menu)



        left_layout.addWidget(self.task_list)

        # 日志区域 - 新增现代化日志记录
        log_group = QGroupBox("📝 执行日志")
        log_layout = QVBoxLayout()

        # 添加清空日志按钮
        log_header_layout = QHBoxLayout()
        log_header_layout.addWidget(QLabel("执行日志:"))
        log_header_layout.addStretch()
        self.clear_log_btn = QPushButton("清空日志")
        self.clear_log_btn.setFixedSize(80, 24)
        self.clear_log_btn.clicked.connect(self.clear_log)
        log_header_layout.addWidget(self.clear_log_btn)

        log_layout.addLayout(log_header_layout)

        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(200)
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        left_layout.addWidget(log_group)

        # 右侧配置区域
        right_panel = QFrame()
        right_panel.setFrameShape(QFrame.StyledPanel)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(15, 15, 15, 15)

        # 任务信息组
        task_info_group = QGroupBox("ℹ️ 任务信息")
        task_info_layout = QFormLayout()
        task_info_layout.setLabelAlignment(Qt.AlignRight)
        task_info_layout.setSpacing(10)

        self.task_name = QLineEdit()
        self.task_name.setPlaceholderText("输入任务名称")
        self.task_name.setFont(QFont("Arial", 10))

        self.task_status = QLabel("未运行")

        task_info_layout.addRow("任务名称:", self.task_name)
        task_info_layout.addRow("当前状态:", self.task_status)

        task_info_group.setLayout(task_info_layout)
        # 定时设置组
        schedule_group = QGroupBox("⏰ 定时设置")
        schedule_layout = QGridLayout()
        schedule_layout.setSpacing(10)
        schedule_layout.setColumnStretch(5, 1)  # 添加弹性空间

        # 执行方式
        schedule_layout.addWidget(QLabel("执行方式:"), 0, 0)
        self.schedule_enable = QComboBox()
        self.schedule_enable.addItems(["立即执行", "定时执行"])
        self.schedule_enable.setMinimumWidth(120)
        self.schedule_enable.currentTextChanged.connect(self.on_schedule_mode_changed)
        schedule_layout.addWidget(self.schedule_enable, 0, 1)

        # 执行时间 - 支持鼠标滚轮
        schedule_layout.addWidget(QLabel("执行时间:"), 0, 2)
        time_widget = QWidget()
        time_layout = QHBoxLayout(time_widget)
        time_layout.setContentsMargins(0, 0, 0, 0)
        time_layout.setSpacing(5)

        self.schedule_time = WheelTimeEdit(QTime.currentTime().addSecs(300))  # 自定义支持滚轮的TimeEdit
        self.schedule_time.setDisplayFormat("HH:mm:ss")
        self.schedule_time.setMinimumWidth(100)
        self.schedule_time.setMaximumWidth(100)
        self.schedule_time.setTimeRange(QTime(0, 0, 0), QTime(23, 59, 59))
        self.schedule_time.setToolTip("使用鼠标滚轮调整时间\n单击可分别编辑时、分、秒")
        time_layout.addWidget(self.schedule_time)

        # 时间快捷按钮
        time_buttons = []
        time_presets = [
            ("13:14", (13, 14)),  # 13:14 时间
            ("晚安时间", (0, 0))  # 0点时间
        ]

        for text, time_values in time_presets:
            btn = QPushButton(text)
            btn.setFixedSize(60, 25)
            btn.setStyleSheet("""
                QPushButton { 
                    font-size: 10px; 
                    padding: 2px; 
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
            """)
            # 根据按钮文本设置不同的点击行为
            if text == "13:14":
                btn.clicked.connect(lambda checked, h=time_values[0], m=time_values[1]: self.set_time_to(h, m))
            elif text == "晚安时间":
                btn.clicked.connect(lambda checked, h=time_values[0], m=time_values[1]: self.set_time_to(h, m))
            time_buttons.append(btn)
            time_layout.addWidget(btn)

        # 需要在类中添加以下方法


        # time_layout.addStretch()
        schedule_layout.addWidget(time_widget, 0, 3, 1, 2)

        # 重复间隔 - 支持鼠标滚轮
        schedule_layout.addWidget(QLabel("重复间隔:"), 1, 0)
        interval_widget = QWidget()
        interval_layout = QHBoxLayout(interval_widget)
        interval_layout.setContentsMargins(0, 0, 0, 0)
        interval_layout.setSpacing(5)

        self.repeat_interval = WheelSpinBox()  # 自定义支持滚轮的SpinBox
        self.repeat_interval.setRange(0, 1440)
        self.repeat_interval.setValue(0)
        self.repeat_interval.setMinimumWidth(80)
        self.repeat_interval.setMaximumWidth(80)
        self.repeat_interval.setSuffix(" 分钟")
        self.repeat_interval.setSpecialValueText("")
        self.repeat_interval.setToolTip("使用鼠标滚轮调整间隔\n")
        self.repeat_interval.valueChanged.connect(self.update_next_run_time)
        interval_layout.addWidget(self.repeat_interval)
        # 间隔快捷按钮
        interval_buttons = []
        interval_presets = [
            ("0分钟", 0), ("24小时", 1440)
        ]

        for text, interval in interval_presets:
            btn = QPushButton(text)
            btn.setFixedSize(55, 25)
            btn.setStyleSheet("""
                QPushButton { 
                    font-size: 10px; 
                    padding: 2px; 
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
            """)
            btn.clicked.connect(lambda checked, i=interval: self.repeat_interval.setValue(i))
            interval_buttons.append(btn)
            interval_layout.addWidget(btn)

        interval_layout.addStretch()
        schedule_layout.addWidget(interval_widget, 1, 1, 1, 2)

        # 重复次数
        schedule_layout.addWidget(QLabel("重复次数:"), 1, 3)
        self.repeat_count = QComboBox()
        self.repeat_count.setEditable(True)  # 设置为可编辑
        self.repeat_count.addItems(["1", "3", "7","9", "无限"])
        self.repeat_count.setCurrentIndex(0)
        self.repeat_count.setMinimumWidth(80)
        self.repeat_count.currentTextChanged.connect(self.update_next_run_time)
        # 添加输入验证器，只允许输入数字或"无限"
        validator = QIntValidator(1, 999999)  # 允许输入1到999999的整数
        self.repeat_count.setValidator(validator)
        self.repeat_count.editTextChanged.connect(self.on_repeat_count_edited)

        schedule_layout.addWidget(self.repeat_count, 1, 4)

        # 下一次执行时间显示
        self.next_run_label = QLabel("下次执行: -")
        self.next_run_label.setStyleSheet("""
            QLabel {
                color: #2c5aa0; 
                font-size: 11px; 
                padding: 8px;
                background-color: #f0f8ff;
                border-radius: 5px;
                border: 1px solid #d0e0f0;
                margin: 2px;
            }
        """)
        self.next_run_label.setMinimumWidth(200)
        self.next_run_label.setAlignment(Qt.AlignCenter)
        self.next_run_label.setWordWrap(True)
        schedule_layout.addWidget(self.next_run_label, 0, 5, 2, 5)

        # 连接信号
        self.schedule_time.timeChanged.connect(self.update_next_run_time)
        self.repeat_interval.valueChanged.connect(self.update_next_run_time)

        schedule_group.setLayout(schedule_layout)

        # 初始化状态
        self.update_next_run_time()
        self.on_schedule_mode_changed("立即执行")
        # 步骤配置区域
        steps_group = QGroupBox("⚙️ 操作步骤配置")
        steps_layout = QVBoxLayout()
        steps_layout.setSpacing(10)

        # 步骤表格 - 设置列宽可拖拽
        self.steps_table = QTableWidget(0, 4)
        self.steps_table.setHorizontalHeaderLabels(["类型", "描述", "参数"])
        self.steps_table.setColumnCount(3)  # 设置列数为3
        self.steps_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)  # 可拖拽调整列宽
        self.steps_table.horizontalHeader().setStretchLastSection(True)
        self.steps_table.verticalHeader().setVisible(False)
        self.steps_table.setSelectionBehavior(QAbstractItemView.SelectRows)

        # 步骤操作按钮 - 添加emoji和快捷键
        step_btn_layout = QHBoxLayout()
        self.add_step_btn = QPushButton("➕ 添加步骤 (A)")
        self.add_step_btn.setShortcut(QKeySequence("Ctrl+A"))
        self.edit_step_btn = QPushButton("✏️ 编辑步骤 (E)")
        self.edit_step_btn.setShortcut(QKeySequence("Ctrl+E"))
        self.remove_step_btn = QPushButton("➖ 删除步骤 (Del)")
        self.remove_step_btn.setShortcut(QKeySequence.Delete)  # 确保删除按钮的快捷键为 Delete
        self.copy_step_btn = QPushButton("📋 复制步骤")
        self.copy_step_btn.setShortcut(QKeySequence("Ctrl+C"))  # 新增：设置复制按钮的快捷键为 Ctrl+C
        self.move_up_btn = QPushButton("⬆️ 上移 (↑)")
        self.move_up_btn.setShortcut(QKeySequence("Ctrl+Up"))
        self.move_down_btn = QPushButton("⬇️ 下移 (↓)")
        self.move_down_btn.setShortcut(QKeySequence("Ctrl+Down"))

        step_btn_layout.addWidget(self.add_step_btn)
        step_btn_layout.addWidget(self.edit_step_btn)
        step_btn_layout.addWidget(self.copy_step_btn)
        step_btn_layout.addWidget(self.remove_step_btn)
        step_btn_layout.addStretch()
        step_btn_layout.addWidget(self.move_up_btn)
        step_btn_layout.addWidget(self.move_down_btn)


        steps_layout.addWidget(self.steps_table)
        steps_layout.addLayout(step_btn_layout)
        steps_group.setLayout(steps_layout)

        # 操作按钮组 - 添加emoji
        action_btn_layout = QHBoxLayout()
        self.start_current_btn = QPushButton("▶️ 开始当前任务")
        self.stop_current_btn = QPushButton("⏹️ 停止当前任务")
        self.stop_current_btn.setEnabled(False)
        self.save_btn = QPushButton("💾 保存配置")

        action_btn_layout.addWidget(self.start_current_btn)
        action_btn_layout.addWidget(self.stop_current_btn)
        action_btn_layout.addStretch()
        action_btn_layout.addWidget(self.save_btn)

        # 添加到右侧布局
        right_layout.addWidget(task_info_group)
        right_layout.addWidget(schedule_group)
        right_layout.addWidget(steps_group)
        right_layout.addLayout(action_btn_layout)

        # 添加左右面板到分割器
        self.splitter.addWidget(left_panel)
        self.splitter.addWidget(right_panel)

        # 恢复分割器位置
        splitter_sizes = self.settings.value("splitterSizes")
        if splitter_sizes:
            splitter_sizes = [int(s) for s in splitter_sizes]
            self.splitter.setSizes(splitter_sizes)
        else:
            self.splitter.setSizes([280, 700])

        # 添加日志区域可拖拽调整高度
        self.log_splitter = QSplitter(Qt.Vertical)
        self.log_splitter.addWidget(self.task_list)
        self.log_splitter.addWidget(log_group)
        left_layout.insertWidget(2, self.log_splitter)
        self.log_splitter.setSizes([300, 150])

        # 添加到主布局
        main_layout.addWidget(self.splitter)

        self.setCentralWidget(main_widget)

        # 创建菜单栏
        self.create_menus()

        # 连接信号
        self.task_list.currentItemChanged.connect(self.task_selected)
        self.new_task_btn.clicked.connect(self.create_new_task)
        self.start_current_btn.clicked.connect(self.start_current_task)
        self.stop_current_btn.clicked.connect(self.stop_current_task)
        self.add_step_btn.clicked.connect(self.add_step)
        self.edit_step_btn.clicked.connect(self.edit_step)
        self.remove_step_btn.clicked.connect(self.remove_step)
        self.move_up_btn.clicked.connect(self.move_step_up)
        self.move_down_btn.clicked.connect(self.move_step_down)
        self.save_btn.clicked.connect(self.save_task_config)
        # self.apply_schedule_btn.clicked.connect(self.apply_schedule)
        self.copy_step_btn.clicked.connect(self.copy_step)

        # 应用当前主题
        self.apply_theme(self.current_theme)

        # 检测系统主题
        self.detect_system_theme()

        # 创建系统托盘图标
        self.create_system_tray()

        # 添加已有任务
        self.load_all_configs("config")
    def set_time_to(self, hour, minute):
        """设置时间为指定的小时和分钟"""
        current_time = QTime(hour, minute, 0)
        self.schedule_time.setTime(current_time)
    def on_schedule_mode_changed(self, mode):
        """执行方式改变时的处理"""
        is_scheduled = mode == "定时执行"

        # # 启用/禁用相关控件
        # self.schedule_time.setEnabled(is_scheduled)
        # self.repeat_interval.setEnabled(is_scheduled)
        # self.repeat_count.setEnabled(is_scheduled)
        #
        # # 启用/禁用快捷按钮
        # for widget in self.findChildren(QPushButton):
        #     if widget.text().endswith('m'):
        #         widget.setEnabled(is_scheduled)

        # 更新提示
        if is_scheduled:
            self.update_next_run_time()
        else:
            self.next_run_label.setText("立即执行模式")
            self.next_run_label.setStyleSheet("""
                QLabel {
                    color: #666; 
                    font-size: 11px; 
                    padding: 5px;
                    background-color: #f8f8f8;
                    border-radius: 3px;
                    border: 1px solid #e0e0e0;
                }
            """)


    # 在类中添加处理编辑的函数
    def on_repeat_count_edited(self, text):
        """处理重复次数编辑事件"""
        # 如果用户输入了"无限"，则设置为"无限"
        if text == "无限":
            return

        # 如果输入的是数字，验证范围
        if text.isdigit():
            value = int(text)
            if value < 1:
                # 如果小于1，设置为1
                self.repeat_count.setCurrentText("1")
            elif value > 999999:
                # 如果大于999999，设置为999999
                self.repeat_count.setCurrentText("999999")
        elif text != "":
            # 如果输入的不是数字也不是"无限"，清除输入
            cursor_pos = self.repeat_count.lineEdit().cursorPosition()
            self.repeat_count.setCurrentText("".join(filter(str.isdigit, text)))
            self.repeat_count.lineEdit().setCursorPosition(min(cursor_pos, len(self.repeat_count.currentText())))

    # 修改获取重复次数值的方法
    def get_repeat_count_value(self):
        """获取重复次数的实际值"""
        text = self.repeat_count.currentText()
        if text == "无限":
            return "无限"
        elif text.isdigit():
            return text
        else:
            return "1"  # 默认值

    def update_next_run_time(self):
        """更新下一次执行时间显示"""
        schedule_type = self.schedule_enable.currentText()

        # 获取当前设置的值
        interval = self.repeat_interval.value()
        repeat_type = self.repeat_count.currentText()

        # 如果是定时执行模式
        if schedule_type == "定时执行":
            schedule_time = self.schedule_time.time()
            now = QTime.currentTime()
            current_date = QDate.currentDate()
            next_run = QTime(schedule_time.hour(), schedule_time.minute(), schedule_time.second())

            # 计算下一次执行日期时间
            if next_run < now:
                # 如果今天的时间已过，则明天执行
                next_date = current_date.addDays(1)
            else:
                next_date = current_date

            next_run_datetime = QDateTime(next_date, next_run)
            next_run_str = next_run_datetime.toString("yyyy-MM-dd HH:mm:ss")

            if interval > 0:
                if repeat_type == "无限":
                    message = f"下次执行: {next_run_str}\n每 {interval} 分钟重复，无限次"
                    color = "#2c5aa0"
                else:
                    message = f"下次执行: {next_run_str}\n每 {interval} 分钟重复，共 {repeat_type} 次"
                    color = "#2c5aa0"
            else:
                message = f"下次执行: {next_run_str}\n无间隔时间 共 {repeat_type} 次"
                color = "#2c5aa0"

            self.next_run_label.setText(message)
            self.next_run_label.setStyleSheet(f"""
                QLabel {{
                    color: {color}; 
                    font-size: 11px; 
                    padding: 8px;
                    background-color: #f0f8ff;
                    border-radius: 5px;
                    border: 1px solid #d0e0f0;
                    margin: 2px;
                }}
            """)
        else:
            # 立即执行模式
            now = QDateTime.currentDateTime()
            next_run_datetime = now.addSecs(1)  # 立即执行，下一次执行时间就是现在
            next_run_datetime_str = next_run_datetime.toString("HH:mm:ss")

            if interval > 0:
                if repeat_type == "无限":
                    message = f"立即执行\n每 {interval} 分钟重复，无限次"
                    color = "#2c5aa0"
                else:
                    message = f"立即执行\n每 {interval} 分钟重复，共 {repeat_type} 次"
                    color = "#2c5aa0"
            else:
                message = f"立即执行，无间隔，共 {repeat_type} 次"
                color = "#2c5aa0"

            self.next_run_label.setText(message)
            self.next_run_label.setStyleSheet(f"""
                QLabel {{
                    color: {color}; 
                    font-size: 11px; 
                    padding: 8px;
                    background-color: #f0f8ff;
                    border-radius: 5px;
                    border: 1px solid #d0e0f0;
                    margin: 2px;
                }}
            """)

    def validate_schedule_settings(self):
        """验证定时设置是否有效"""
        interval = self.repeat_interval.value()

        if interval < 0 or interval > 1440:
            QMessageBox.warning(self, "无效设置", "重复间隔必须在0-1440分钟之间")
            return False

        schedule_time = self.schedule_time.time()
        if not schedule_time.isValid():
            QMessageBox.warning(self, "无效时间", "请选择有效的执行时间")
            return False

        return True
    def setup_hotkey_listener(self):
        """启动 Esc 热键监听"""
        self.hotkey_listener = HotkeyListener(self)
        self.hotkey_listener.hotkey_activated.connect(self.on_esc_pressed)
        self.hotkey_listener.start()  # 启动线程
    @Slot()
    def on_esc_pressed(self):
        """响应 Esc 键（在主线程执行）"""
        if self.task_runner and self.task_thread and self.task_thread.is_alive():
            self.stop_current_task()
            self.statusBar().showMessage("🛑 Esc 被按下，任务已停止", 2000)


    def load_all_configs(self, config_dir="config"):
        """
        扫描 config_dir 内所有 *.json 并加载为任务
        """
        # 获取配置目录的绝对路径
        if getattr(sys, 'frozen', False):
            # 打包后的情况：在可执行文件同级目录下查找 config
            application_path = os.path.dirname(sys.executable)
            config_dir = os.path.join(application_path, config_dir)
        else:
            # 开发环境
            config_dir = resource_path(config_dir)
        if not os.path.isdir(config_dir):
            os.makedirs(config_dir, exist_ok=True)
            return
        first_task_loaded = False  # 标记是否已加载第一个任务
        for fname in os.listdir(config_dir):
            if not fname.endswith(".json"):
                continue
            path = os.path.join(config_dir, fname)
            try:
                if path:
                    try:
                        with open(path, 'r') as f:
                            task_config = json.load(f)

                        task_name = task_config.get("name", False)
                        if task_name:
                            self.add_task(task_name)
                            self.tasks[task_name] = task_config

                        # 选中新导入的任务
                        # 如果是第一个加载的任务，则选中并显示其配置
                        if not first_task_loaded:
                            for i in range(self.task_list.count()):
                                item = self.task_list.item(i)
                                widget = self.task_list.itemWidget(item)

                                if widget and widget.task_name == task_name:
                                    self.task_list.setCurrentItem(item)
                                    self.display_task_config(task_name)
                                    first_task_loaded = True
                                    break
                    except Exception as e:
                        QMessageBox.critical(self, "导入失败", f"导入配置时出错: {str(e)}")
            except Exception as e:
                print(f"加载配置 {path} 失败：{e}")

    def display_task_config(self, task_name):
        """
        显示指定任务的配置数据到定时设置和操作步骤配置区域
        """
        if task_name not in self.tasks:
            return

        task_config = self.tasks[task_name]

        # 显示任务名称
        self.task_name.setText(task_name)

        # 显示定时设置
        schedule = task_config.get("schedule", {})
        self.schedule_enable.setCurrentText(schedule.get("enable", "立即执行"))
        time_str = schedule.get("time", QTime.currentTime().toString("HH:mm:ss"))

        # 解析时间字符串
        time_parts = time_str.split(":")
        if len(time_parts) == 3:
            hour, minute, second = map(int, time_parts)
            self.schedule_time.setTime(QTime(hour, minute, second))

        self.repeat_interval.setValue(int(schedule.get("interval", 0)))
        self.repeat_count.setCurrentText(str(schedule.get("repeat", "1")))

        # 显示步骤配置
        steps = task_config.get("steps", [])
        self.steps_table.setRowCount(0)  # 清空现有步骤

        for step in steps:
            self.add_step_to_table(step)
    def create_system_tray(self):
        """创建系统托盘图标"""
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = QAction("显示窗口", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)

        hide_action = QAction("隐藏窗口", self)
        hide_action.triggered.connect(self.hide)
        tray_menu.addAction(hide_action)

        tray_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self.close)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_icon_activated)
        self.tray_icon.show()

        # 连接信号
        self.tray_icon.messageClicked.connect(self.tray_message_clicked)

    def tray_icon_activated(self, reason):
        """托盘图标被激活时的处理"""
        if reason == QSystemTrayIcon.DoubleClick:
            if self.isVisible():
                self.hide()
            else:
                self.showNormal()
                self.activateWindow()

    def tray_message_clicked(self):
        """托盘消息被点击时的处理"""
        self.showNormal()
        self.activateWindow()

    def closeEvent(self, event):
        """重写关闭事件，实现最小化到托盘"""
        if self.tray_icon.isVisible():
            self.hide()
            event.ignore()
        else:
            # 保存设置
            self.save_settings()
            # 停止所有定时器
            for timer in self.scheduled_timers.values():
                timer.stop()
            event.accept()

    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
        self.log_text.appendPlainText(f"[{time.strftime('%H:%M:%S')}] 日志已清空")

    def show_ai_token_config(self):
        """显示AI Token配置对话框"""
        dialog = AITokenConfigDialog(self)
        dialog.exec()

    # 在主窗口类中添加以下方法

    def show_ai_test(self):
        """显示 AI 测试对话框"""
        try:
            dialog = AITestDialog(self)
            dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开 AI 测试对话框: {str(e)}")

    def create_menus(self):
        menu_bar = self.menuBar()

        # === 新增：设置菜单 ===
        settings_menu = menu_bar.addMenu("⚙️ 设置")
        # 主容器
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)
        settings_layout.setContentsMargins(8, 4, 8, 4)
        settings_layout.setSpacing(6)

        # 1. 自动跳过 + 超时时间（纵向）
        # --- 自动跳过复选框 ---
        self.auto_skip_checkbox = QCheckBox("图片查找超时后自动跳过")
        self.auto_skip_checkbox.setChecked(False)

        # --- 超时时间（水平布局）---
        timeout_layout = QHBoxLayout()
        timeout_label = QLabel("超时时间:")
        self.timeout_spinbox = QDoubleSpinBox()
        self.timeout_spinbox.setRange(0, 86400)
        self.timeout_spinbox.setSingleStep(0.5)
        self.timeout_spinbox.setValue(3)
        self.timeout_spinbox.setSuffix(" s")
        self.timeout_spinbox.setFixedWidth(100)
        timeout_layout.addWidget(timeout_label)
        timeout_layout.addWidget(self.timeout_spinbox)
        timeout_layout.addStretch()

        # 2. 鼠标移动设置（水平布局）
        mouse_layout = QHBoxLayout()
        self.instant_click_checkbox = QCheckBox("直接点击")
        self.instant_click_checkbox.setChecked(False)

        self.move_duration_spinbox = QDoubleSpinBox()
        self.move_duration_spinbox.setRange(0.0, 10.0)
        self.move_duration_spinbox.setSingleStep(0.1)
        self.move_duration_spinbox.setValue(0.3)
        self.move_duration_spinbox.setDecimals(1)
        self.move_duration_spinbox.setSuffix(" s")
        self.move_duration_spinbox.setFixedWidth(80)
        self.move_duration_spinbox.setEnabled(True)

        # 3. 窗口最小化设置（新增）
        minimize_layout = QHBoxLayout()
        self.minimize_during_execution_checkbox = QCheckBox("执行任务时最小化窗口")
        self.minimize_during_execution_checkbox.setChecked(True)  # 默认勾选

        minimize_layout.addWidget(self.minimize_during_execution_checkbox)
        minimize_layout.addStretch()

        # 4. label颜色设置（新增）
        label_color_layout = QHBoxLayout()
        self.label_color_checkbox = QCheckBox("开启步骤表格的五彩色")
        self.label_color_checkbox.setChecked(True)  # 默认勾选

        label_color_layout.addWidget(self.label_color_checkbox)
        label_color_layout.addStretch()

        # 连接 checkbox 控制 spinbox 启用状态
        def on_instant_click_toggled(checked):
            self.move_duration_spinbox.setEnabled(not checked)

        self.instant_click_checkbox.toggled.connect(on_instant_click_toggled)

        mouse_layout.addWidget(self.instant_click_checkbox)
        mouse_layout.addWidget(self.move_duration_spinbox)
        mouse_layout.addStretch()

        # 添加到主布局
        settings_layout.addWidget(self.auto_skip_checkbox)
        settings_layout.addLayout(timeout_layout)
        settings_layout.addLayout(mouse_layout)
        settings_layout.addLayout(minimize_layout)  # 添加新行
        settings_layout.addLayout(label_color_layout)  # 添加新行

        # 包装为菜单项
        action = QWidgetAction(settings_menu)
        action.setDefaultWidget(settings_widget)
        settings_menu.addAction(action)

        # === 新增：AI Token 配置菜单项 ===
        ai_token_action = QAction("🤖 AI Token 配置", self)
        ai_token_action.triggered.connect(self.show_ai_token_config)
        settings_menu.addAction(ai_token_action)

        # === 新增：AI 测试菜单项 ===
        ai_test_action = QAction("🧠 AI 测试", self)
        ai_test_action.triggered.connect(self.show_ai_test)
        settings_menu.addAction(ai_test_action)
        # 为设置菜单添加样式
        settings_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 文件菜单
        file_menu = menu_bar.addMenu("📁 文件")
        new_action = QAction("📝 新建任务", self)
        save_action = QAction( "💾 保存配置", self)
        export_action = QAction( "📤 导出配置", self)
        import_action = QAction( "📥 导入配置", self)
        exit_action = QAction( "🚪 退出", self)

        file_menu.addAction(new_action)
        file_menu.addAction(save_action)
        file_menu.addSeparator()
        file_menu.addAction(export_action)
        file_menu.addAction(import_action)
        file_menu.addSeparator()
        file_menu.addAction(exit_action)
        # 关键：给菜单本身设置样式表
        file_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 编辑菜单
        edit_menu = menu_bar.addMenu("✏️ 编辑")
        add_step_action = QAction( "➕ 添加步骤", self)
        edit_step_action = QAction( "✏️ 编辑步骤", self)
        remove_step_action = QAction( "➖ 删除步骤", self)
        copy_step_action = QAction("📋 复制步骤", self)

        edit_menu.addAction(add_step_action)
        edit_menu.addAction(edit_step_action)
        edit_menu.addAction(copy_step_action)
        edit_menu.addAction(remove_step_action)

        edit_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 主题菜单（位于编辑和帮助之间）
        theme_menu = menu_bar.addMenu("🎨 主题")
        theme_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        self.light_theme_action = QAction("☀️ 明亮主题", self)
        self.light_theme_action.setCheckable(True)
        self.light_theme_action.triggered.connect(lambda: self.switch_theme("light"))

        self.dark_theme_action = QAction("🌙 暗黑主题", self)
        self.dark_theme_action.setCheckable(True)
        self.dark_theme_action.triggered.connect(lambda: self.switch_theme("dark"))

        self.system_theme_action = QAction("🔄 跟随系统", self)
        self.system_theme_action.setCheckable(True)
        self.system_theme_action.triggered.connect(lambda: self.switch_theme("system"))

        theme_menu.addAction(self.light_theme_action)
        theme_menu.addAction(self.dark_theme_action)
        theme_menu.addAction(self.system_theme_action)

        # 设置当前主题选中状态
        if self.current_theme == "light":
            self.light_theme_action.setChecked(True)
        elif self.current_theme == "dark":
            self.dark_theme_action.setChecked(True)
        else:
            self.system_theme_action.setChecked(True)

        # 帮助菜单
        help_menu = menu_bar.addMenu("❓ 帮助")
        about_action = QAction("ℹ️ 关于", self)
        docs_action = QAction("📚 使用文档", self)

        help_menu.addAction(docs_action)
        help_menu.addAction(about_action)
        help_menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)

        # 连接菜单信号
        # new_action.triggered.connect(self.create_new_task)
        save_action.triggered.connect(self.save_task_config)
        export_action.triggered.connect(self.export_config)
        import_action.triggered.connect(self.import_config)
        exit_action.triggered.connect(self.close)
        add_step_action.triggered.connect(self.add_step)
        edit_step_action.triggered.connect(self.edit_step)
        remove_step_action.triggered.connect(self.remove_step)
        docs_action.triggered.connect(self.show_docs)
        about_action.triggered.connect(self.show_about)
        copy_step_action.triggered.connect(self.copy_step)

    def show_docs(self):
        """显示使用文档对话框"""
        dialog = QDialog(self)
        dialog.setWindowTitle("使用文档")
        dialog.setMinimumWidth(500)
        dialog.setMinimumHeight(500)

        layout = QVBoxLayout(dialog)

        text = QTextEdit()
        text.setReadOnly(True)
        text.setHtml("""
            <h2>自动化任务管理器使用文档</h2>
            <p>欢迎使用自动化任务管理器！本工具可以帮助您自动化执行重复的计算机操作。</p>
            <p>github开源链接：https://github.com/junior6666/AutoTask-UI-</p

            <h3>基本功能</h3>
            <ul>
                <li><b>创建任务</b>：点击"新建任务"按钮创建新任务</li>
                <li><b>添加步骤</b>：在任务中添加鼠标点击、文本输入、等待等操作步骤</li>
                <li><b>定时执行</b>：根据任务需求设置任务的执行时间，点击开始当前任务按钮即可</li>
                <li><b>执行日志</b>：查看任务执行过程中的详细日志</li>
            </ul>

            <h3>配置说明</h3>
            <p>配置任务时，请确保：</p>
            <ul>
                <li>图片路径正确不含有中文，且图片在屏幕上可见</li>
                <li>设置合适的识别精度和超时时间</li>
                <li>为需要等待的操作添加适当的延时</li>
            </ul>

            <h3>QQ交流群</h3>
            <p>加入我们的QQ交流群获取更多帮助：<b>1057721699</b></p>

            <h3>常见问题</h3>
            <p><b>Q: 为什么找不到图片？</b><br>
            A: 请确保图片在屏幕上可见，且识别精度设置合适（建议0.8-0.9）</p>

            <p><b>Q: 任务执行失败怎么办？</b><br>
            A: 查看执行日志中的错误信息，调整步骤参数后重试</p>
            <p><b>Q: 13:14如何计算的？</b><br>
            A: 无论用户什么时候点击按钮，文案中的“相恋时间”都以 今天 13:14 为截止点计算。以确保定时在13：14发送的逻辑</p>
            <p><b>Q: 开发框架？</b><br>
            A: GUI：🐍 PySide6
            自动化：🤖 PyAutoGUI + 🔍 OpenCV</p>
            <p><b>Q: 开发时长？</b><br>
            A: 核心功能实现 2 days 不过一直在断断续续完善UI和修复各种bug 也欢迎大家参与到源码的开发</p>
            <p><b>Q: pyautogui在定位图片位置时，若屏幕中有两个相同的图片，它会选择哪一个图片？？</b><br>
            A: “谁最靠左上角，谁就中标；后面的即使一模一样也不会被理会。”
如果你想把所有相同图标都找出来，就必须用 locateAllOnScreen()，它会返回一个可迭代对象，里面包含所有匹配区域的坐标盒（left, top, width, height），顺序同样是先上后下、先左后右。（已实现）</p>
        """)

        layout.addWidget(text)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok)
        button_box.accepted.connect(dialog.accept)
        layout.addWidget(button_box)

        dialog.exec()

    def show_about(self):
        # 任意窗口里
        AboutDialog(self).exec()


    def add_task(self, name):
        # 创建自定义列表项
        item = QListWidgetItem(self.task_list)
        item_widget = TaskItemWidget(name, self)
        item.setSizeHint(QSize(0, 45))  # 固定高度确保按钮完全显示
        self.task_list.addItem(item)
        self.task_list.setItemWidget(item, item_widget)

        # 初始化任务配置
        self.tasks[name] = {
            "name": name,
            "schedule": {
                "enable": "立即执行",
                "time": QTime.currentTime().toString("HH:mm:ss"),
                "interval": 0,
                "repeat": "1"
            },
            "steps": []
        }

        # 应用当前主题样式
        self.apply_button_style(item_widget)

        # 选中新添加的任务
        if self.task_list.count() == 1:
            self.task_list.setCurrentItem(item)

    def create_new_task(self):
        name = f"新任务 {self.task_list.count() + 1}"
        self.add_task(name)
        ts = time.strftime("%H:%M:%S")
        # 日志带 emoji
        self.log_text.appendPlainText(f"[{ts}] ✅ [{name}] 已创建！")

    def duplicate_task(self, name):
        new_name = f"{name} 副本"
        self.add_task(new_name)

        # 复制任务配置
        if name in self.tasks:
            self.tasks[new_name] = self.tasks[name].copy()
            self.tasks[new_name]["name"] = new_name
        # 日志带 emoji
        ts = time.strftime("%H:%M:%S")
        self.log_text.appendPlainText(f"[{ts}] 📋 [{name}] → [{new_name}] 已复制！")

    def rename_task(self, name):
        """重命名任务：同步所有内部结构与 UI，保证顺序一致。"""
        if name not in self.tasks:
            return

        # 1. 弹窗获取新名称
        new_name, ok = QInputDialog.getText(
            self, "重命名任务", "请输入新名称：", text=name
        )
        if not ok:
            return
        new_name = new_name.strip()
        if not new_name:
            QMessageBox.warning(self, "提示", "任务名称不能为空！")
            return
        if new_name == name:
            return
        if new_name in self.tasks:
            QMessageBox.warning(self, "提示", f"任务“{new_name}”已存在。")
            return

        # 2. 找到旧 item 的行号
        row = -1
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == name:
                row = i
                break
        if row == -1:  # 理论上不会发生
            return

        # 3. 创建新 item，并插回原位置
        new_item = QListWidgetItem()
        new_widget = TaskItemWidget(new_name, self)
        new_item.setSizeHint(QSize(0, 45))

        # 4. 复制任务数据
        self.tasks[new_name] = self.tasks[name].copy()
        self.tasks[new_name]["name"] = new_name

        # 5. 替换 UI：先插新的，再删旧的
        self.task_list.insertItem(row, new_item)
        self.task_list.setItemWidget(new_item, new_widget)
        self.task_list.takeItem(row + 1)  # 原来的那行现在是 row+1
        if name in self.scheduled_timers:  # 如果之前有时钟，一起迁移
            self.scheduled_timers[new_name] = self.scheduled_timers.pop(name)

        # 6. 选中新任务并保持焦点
        self.task_list.setCurrentItem(new_item)
        self.apply_button_style(new_widget)

        # 7. 彻底删除旧任务
        del self.tasks[name]
        self.on_log_message(name, f"📝 重命名：{name} → {new_name}")

    def delete_task(self, name):
        row = -1
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == name:
                row = i
                break

        if row >= 0:
            self.task_list.takeItem(row)
            if name in self.tasks:
                # 如果任务有定时器，先停止
                if name in self.scheduled_timers:
                    self.scheduled_timers[name].stop()
                    del self.scheduled_timers[name]
                del self.tasks[name]
            self.on_log_message(name, f"🗑️ 已删除任务：{name}")

            # 新增：检查是否删除了最后一个任务
            if self.task_list.count() == 0:
                # 清空当前任务的配置显示
                self.task_name.clear()
                self.task_status.setText("未选择任务")

                # 重置定时设置
                self.schedule_enable.setCurrentIndex(0)  # "立即执行"
                self.schedule_time.setTime(QTime.currentTime())
                self.repeat_interval.setValue(0)
                self.repeat_count.setCurrentIndex(0)  # "1次"

                # 清空步骤表格
                self.steps_table.setRowCount(0)

                # 重置当前任务引用
                self.current_task = None

                # 重置按钮状态
                self.start_current_btn.setEnabled(False)
                self.stop_current_btn.setEnabled(False)

                self.on_log_message("系统", "📋 最后一个任务已删除，配置已重置")

    def task_selected(self, current, previous):
        if current:
            widget = self.task_list.itemWidget(current)
            if widget:
                task_name = widget.task_name
                self.current_task = task_name
                self.task_name.setText(task_name)
                self.task_status.setText(widget.status_label.text())

                # 更新按钮状态
                if widget.status_label.text() == "运行中":
                    self.start_current_btn.setEnabled(False)
                    self.stop_current_btn.setEnabled(True)
                else:
                    self.start_current_btn.setEnabled(True)
                    self.stop_current_btn.setEnabled(False)

                # 加载任务配置
                self.load_task_config(task_name)

    def load_task_config(self, task_name):
        if task_name in self.tasks:
            task_config = self.tasks[task_name]

            # 加载定时设置
            self.schedule_enable.setCurrentText(task_config["schedule"]["enable"])
            self.schedule_time.setTime(QTime.fromString(task_config["schedule"]["time"], "HH:mm:ss"))
            self.repeat_interval.setValue(task_config["schedule"]["interval"])
            self.repeat_count.setCurrentText(task_config["schedule"]["repeat"])

            # 加载步骤
            self.steps_table.setRowCount(0)
            for step in task_config["steps"]:
                self.add_step_to_table(step)

    def add_step_to_table(self, step):
        row = self.steps_table.rowCount()
        self.steps_table.insertRow(row)

        # self.steps_table.setItem(row, 0, QTableWidgetItem(step["type"]))
        # self.steps_table.setItem(row, 1, QTableWidgetItem(StepTableHelper.desc_of(step)))
        use_color = self.label_color_checkbox.isChecked() if hasattr(self, 'label_color_checkbox') else True
        type_widget = StepTableHelper.type_widget(step["type"], use_color)
        self.steps_table.setCellWidget(row, 0, type_widget)
        w = StepTableHelper.widget_of(step,use_color)
        self.steps_table.setCellWidget(row, 1, w)
        self.steps_table.setRowHeight(row, max(StepTableHelper.IMG_HEIGHT + 4, 24))
        self.steps_table.verticalHeader().setDefaultSectionSize(
            StepTableHelper.FIXED_ROW_HEIGHT
        )
        self.steps_table.horizontalHeader().setStretchLastSection(True)

        # 格式化参数显示
        params_text = ""
        if step["type"] == "鼠标点击":
            use_image = step['params'].get('use_image', True)
            use_coordinates = step['params'].get('use_coordinates', False)

            if use_image:
                image_path = step['params'].get('image_path', '')
                image_name = os.path.basename(image_path) if image_path else "未设置"
                click_type = step['params'].get('click_type', '左键单击')
                scan_direction = step['params'].get('scan_direction', '默认')
                offset_x = step['params'].get('offset_x', 0)
                offset_y = step['params'].get('offset_y', 0)
                confidence = step['params'].get('confidence', 0.8)
                timeout = step['params'].get('timeout', 10)

                params_text = f"图片: {image_name}, 点击: {click_type}, 方向: {scan_direction}"
                if offset_x != 0 or offset_y != 0:
                    params_text += f", 偏移: ({offset_x}, {offset_y})"
                params_text += f", 置信度: {confidence}, 超时: {timeout}s"

            elif use_coordinates:
                x_coord = step['params'].get('x_coordinate', 0)
                y_coord = step['params'].get('y_coordinate', 0)
                click_type = step['params'].get('click_type', '左键单击')
                offset_x = step['params'].get('offset_x', 0)
                offset_y = step['params'].get('offset_y', 0)

                params_text = f"坐标: ({x_coord}, {y_coord}), 点击: {click_type}"
                if offset_x != 0 or offset_y != 0:
                    params_text += f", 偏移: ({offset_x}, {offset_y})"

            else:
                params_text = "未启用图片或坐标模式"
        elif step["type"] == "文本输入":
            params_text = f"文本: {step['params'].get('text', 'excel表内容')}"
        elif step["type"] == "等待":
            params_text = f"等待: {step['params'].get('seconds', 0)}秒"
        elif step["type"] == "截图":
            params_text = f"保存到: {step['params'].get('save_path', '')}"
        elif step["type"] == "鼠标滚轮":
            params_text = f"鼠标滚轮: {step['params'].get('direction', '向下滚动')},{step['params'].get('clicks', '3')}格"
        elif step["type"] == "键盘热键":
            hotkey = step["params"].get("hotkey", "ctrl+c").upper()
            delay = step["params"].get("delay_ms", 100)
            params_text = f"键盘热键: {hotkey}, 延时 {delay} ms"
        elif step["type"] == "拖拽":
            use_image = step['params'].get('use_image', True)
            if use_image:
                img_path = step['params'].get('image_path', '')
                if img_path:
                    img_name = os.path.basename(img_path)
                    dx = step['params'].get('drag_x', 0)
                    dy = step['params'].get('drag_y', 0)
                    params_text = f"图片: {img_name} (横向距离{dx},纵向距离{dy})"
                else:
                    params_text = "图片: 未设置"
            else:
                start_x = step['params'].get('start_x', 0)
                start_y = step['params'].get('start_y', 0)
                end_x = step['params'].get('end_x', 0)
                end_y = step['params'].get('end_y', 0)
                params_text = f"从({start_x},{start_y})到({end_x},{end_y})"
        elif step["type"] == "自动回复":
            provider = step['params'].get("provider", "kimi")
            system_prompt = step['params'].get("system_prompt", "")
            use_history = step['params'].get("use_history", True)
            stream = step['params'].get("stream", False)

            # 构造参数描述文本
            params_text = (
                f"AI 提供商: {provider}, "
                f"系统提示词: {'已设置' if system_prompt else '未设置'}, "
                f"使用历史记录: {'是' if use_history else '否'}, "
                f"流式输出: {'是' if stream else '否'}"
            )

        self.steps_table.setItem(row, 2, QTableWidgetItem(params_text))
        # self.steps_table.setItem(row, 3, QTableWidgetItem(str(step.get("delay", 0))))
        self.steps_table.resizeColumnToContents(1)  # 列宽按内容自适应

    def start_current_task(self):
        if not self.current_task:
            return

        # 检查是否有定时设置
        schedule_type = self.schedule_enable.currentText()
        if schedule_type != "立即执行":
            # 验证定时设置
            if not self.validate_schedule_settings():
                return
            # 处理定时执行逻辑
            task_name = self.current_task

            # 如果任务已有定时器，先停止
            if task_name in self.scheduled_timers:
                self.scheduled_timers[task_name].stop()
                del self.scheduled_timers[task_name]

            # 获取定时设置
            schedule_time = self.schedule_time.time()

            # 计算第一次执行的时间
            now = QTime.currentTime()
            first_run = QTime(schedule_time.hour(), schedule_time.minute(), schedule_time.second()).addSecs(-10)

            # 如果当前时间已超过设定时间，则明天执行
            if first_run < now:
                first_run = first_run.addSecs(24 * 3600)  # 加一天

            # 计算延迟时间（毫秒）
            delay_ms = now.msecsTo(first_run)

            # 更新主界面按钮状态
            self.start_current_btn.setEnabled(False)
            self.stop_current_btn.setEnabled(True)

            # 更新任务列表中的状态（只更新当前任务）
            for i in range(self.task_list.count()):
                item = self.task_list.item(i)
                widget = self.task_list.itemWidget(item)
                if widget and widget.task_name == self.current_task:
                    widget.status_label.setText("定时执行中")
                    widget.start_btn.setEnabled(False)
                    widget.stop_btn.setEnabled(True)
                    break

            # 创建首次执行的定时器
            initial_timer = QTimer(self)
            initial_timer.setSingleShot(True)  # 只执行一次

            def run_initial_task():
                # 执行倒计时并运行任务
                # 将重复间隔和重复次数传递给任务执行函数
                self.run_task_with_countdown(task_name)

            initial_timer.timeout.connect(run_initial_task)
            initial_timer.start(delay_ms)

            # 保存定时器引用
            self.scheduled_timers[task_name] = initial_timer

            # 显示提示信息
            first_run_1 = first_run.addSecs(10)
            first_run_str = first_run_1.toString('HH:mm:ss')
            self.log_text.appendPlainText(
                f"[{time.strftime('%H:%M:%S')}] 已设置定时任务: {task_name} 将在 {first_run_str} 执行")

            # 显示状态栏信息（不修改全局状态，只显示当前设置信息）
            self.statusBar().showMessage(f"定时任务已设置，将在 {first_run_str} 执行 {task_name}")

            QMessageBox.information(self, "定时成功",
                                    f"[{time.strftime('%H:%M:%S')}] 已设置定时任务: {task_name} 将在 {first_run_str} 执行\n请保持桌面处于从不熄屏状态")

            return  # 如果是定时执行，直接返回，不立即执行任务
        # 立即执行任务的逻辑
        elif schedule_type == "立即执行":
            self.execute_task_immediately(self.current_task)

    def run_task_with_countdown(self, task_name,countdown_seconds = 10):
        """执行带倒计时的任务"""
        # 创建倒计时定时器
        countdown_timer = QTimer(self)
        countdown_timer.setInterval(1000)  # 每秒触发一次
        def update_countdown():
            nonlocal countdown_seconds
            current_time = time.strftime('%H:%M:%S')  # 获取当前时间
            if countdown_seconds > 0:
                self.statusBar().showMessage(
                    f"[{current_time}] 任务 '{task_name}' 即将执行: {countdown_seconds}秒"
                )
                countdown_seconds -= 1
            else:
                countdown_timer.stop()
                current_time = time.strftime('%H:%M:%S')  # 再次获取当前时间
                self.statusBar().showMessage(
                    f"[{current_time}] 任务 '{task_name}' 开始执行"
                )
                # 实际执行任务
                self.execute_task_immediately(task_name)
        # 启动倒计时
        countdown_timer.timeout.connect(update_countdown)
        countdown_timer.start()
        # 立即更新一次倒计时显示
        current_time = time.strftime('%H:%M:%S')
        self.statusBar().showMessage(
            f"[{current_time}] 任务 '{task_name}' 即将执行: {countdown_seconds}秒"
        )
        # 保存倒计时定时器引用以便可以停止
        if not hasattr(self, 'countdown_timers'):
            self.countdown_timers = {}
        self.countdown_timers[task_name] = countdown_timer
    def execute_task_immediately(self,task_name):
        """立即执行任务的公共方法"""
        # if task_name not in self.current_task:
        #     return

        # 清除状态栏的倒计时信息
        self.statusBar().showMessage("")

        # 获取任务配置
        task_config = self.tasks.get(task_name, {})
        steps = task_config.get("steps", [])

        if not steps:
            QMessageBox.warning(self, "无法启动", "当前任务没有配置任何步骤")
            return

        auto_skip = self.auto_skip_checkbox.isChecked()  # ✅ 读取 QCheckBox 状态
        timeout = self.timeout_spinbox.value()  # 获取用户设置的超时时间
        instant_click = self.instant_click_checkbox.isChecked()
        move_duration = self.move_duration_spinbox.value() if not instant_click else 0.0

        # 创建任务运行器
        self.task_runner = TaskRunner(task_name, steps,
                                      auto_skip_image_timeout=auto_skip,
                                      timeout=timeout,
                                      instant_click=instant_click,
                                      move_duration=move_duration,
                                      parent=self)

        # 设置重复次数
        repeat_text = self.repeat_count.currentText()

        if self.repeat_interval.value() == 0:
            if repeat_text == "无限":
                self.task_runner.set_repeat_count(99999)  # 设置一个很大的数表示无限
            else:
                count = int(repeat_text)
                self.task_runner.set_repeat_count(count)
        elif self.repeat_interval.value() > 0:
            self.task_runner.set_repeat_interval(self.repeat_interval.value())
            if repeat_text == "无限":
                self.task_runner.set_repeat_count(99999)  # 设置一个很大的数表示无限
            else:
                count = int(repeat_text)
                self.task_runner.set_repeat_count(count)
        # 连接信号
        self.task_runner.task_completed.connect(self.on_task_completed)
        self.task_runner.task_progress.connect(self.on_task_progress)
        self.task_runner.log_message.connect(self.on_log_message)  # 连接日志信号

        # 在单独线程中运行任务
        self.task_thread = threading.Thread(target=self.task_runner.run)
        self.task_thread.daemon = True
        self.task_thread.start()

        # 更新UI状态
        self.start_current_btn.setEnabled(False)
        self.stop_current_btn.setEnabled(True)
        self.task_status.setText("运行中")

        # 更新任务列表中的状态
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == self.current_task:
                widget.status_label.setText("运行中")
                widget.start_btn.setEnabled(False)
                widget.stop_btn.setEnabled(True)
                break
        if hasattr(self, 'minimize_during_execution_checkbox') and \
                self.minimize_during_execution_checkbox.isChecked():
        # 新增：任务开始后最小化窗口
            self.showMinimized()
        # self.statusBar().showMessage("任务执行完成")


    def stop_current_task(self):
        # 停止当前运行的任务
        if self.task_runner and self.task_runner.is_running:
            self.task_runner.stop()

        # 停止当前任务的定时器（如果有）
        if self.current_task and self.current_task in self.scheduled_timers:
            timer = self.scheduled_timers[self.current_task]
            if timer and timer.isActive():
                timer.stop()
            del self.scheduled_timers[self.current_task]

        # 停止当前任务的倒计时（如果有）
        if hasattr(self, 'countdown_timers') and self.current_task in self.countdown_timers:
            countdown_timer = self.countdown_timers[self.current_task]
            if countdown_timer and countdown_timer.isActive():
                countdown_timer.stop()
            del self.countdown_timers[self.current_task]

            # 记录日志
            self.log_text.appendPlainText(
                f"[{time.strftime('%H:%M:%S')}] 已取消定时任务: {self.current_task}")

        # 更新UI状态
        self.start_current_btn.setEnabled(True)
        self.stop_current_btn.setEnabled(False)
        self.task_status.setText("已停止")

        # 清除状态栏的倒计时信息
        self.statusBar().showMessage("任务已停止")

        # 更新任务列表中的状态
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == self.current_task:
                widget.status_label.setText("已停止")
                widget.start_btn.setEnabled(True)
                widget.stop_btn.setEnabled(False)
                break

        # 恢复窗口显示（如果之前最小化了）
        self.showNormal()
    def cleanup_scheduled_timers(self):
        """清理无效的定时器"""
        tasks_to_remove = []
        for task_name, timer in self.scheduled_timers.items():
            if timer is None or not timer.isActive():
                tasks_to_remove.append(task_name)

        for task_name in tasks_to_remove:
            del self.scheduled_timers[task_name]

        if tasks_to_remove:
            self.log_text.appendPlainText(
                f"[{time.strftime('%H:%M:%S')}] 清理了 {len(tasks_to_remove)} 个无效定时器")
    def stop_all_scheduled_tasks(self):
        """停止所有定时任务"""
        tasks_stopped = []
        for task_name, timer in list(self.scheduled_timers.items()):
            if timer and timer.isActive():
                timer.stop()
                tasks_stopped.append(task_name)

        # 清空定时器字典
        self.scheduled_timers.clear()

        # 记录日志
        if tasks_stopped:
            self.log_text.appendPlainText(
                f"[{time.strftime('%H:%M:%S')}] 已停止所有定时任务: {', '.join(tasks_stopped)}")
            self.statusBar().showMessage(f"已停止 {len(tasks_stopped)} 个定时任务")
    def closeEvent(self, event):
        # 清理热键监听
        if self.hotkey_listener and self.hotkey_listener.isRunning():
            self.hotkey_listener.stop()

        super().closeEvent(event)

    def on_task_completed(self, task_name, success, message):
        # 新增：任务完成后恢复窗口显示
        self.showNormal()

        # 更新UI状态
        self.start_current_btn.setEnabled(True)
        self.stop_current_btn.setEnabled(False)
        self.task_status.setText("已停止" if success else "已中断")

        # 更新任务列表中的状态
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget and widget.task_name == task_name:
                widget.status_label.setText("已停止" if success else "已中断")
                widget.start_btn.setEnabled(True)
                widget.stop_btn.setEnabled(False)
                break

        # 记录日志
        # self.log_text.appendPlainText(f"[{time.strftime('%H:%M:%S')}] {message}")

    def on_task_progress(self, task_name, current, total):
        self.task_status.setText(f"运行中 ({current}/{total})")

    def on_log_message(self, task_name, message):
        """处理日志消息"""
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        log_entry = f"[{timestamp}] [{task_name}] {message}"
        self.log_text.appendPlainText(log_entry)

        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )

    def show_context_menu(self, pos):
        # 获取点击位置的item
        item = self.task_list.itemAt(pos)
        if not item:
            return

        # 创建上下文菜单
        menu = QMenu(self)

        # 获取任务名称
        widget = self.task_list.itemWidget(item)
        task_name = widget.task_name if widget else ""

        # 添加菜单项
        rename_action = menu.addAction("✏️ 重命名")
        duplicate_action = menu.addAction("📋 创建副本")
        menu.addSeparator()
        delete_action = menu.addAction("🗑️ 删除任务")
        menu.setStyleSheet("""
              QMenu {
                  /* 可选：菜单整体背景 */
                  background: #ffffff;
                  border: 1px solid #cccccc;
              }

              QMenu::item {
                  /* 普通状态下的文字背景 */
                  background: transparent;
                  padding: 6px 20px;
                  color: black;
              }

              QMenu::item:selected {       /* 鼠标悬停/键盘选中时生效 */
                  background: #dbeafe;     /* 你想要的 hover 背景色 */
                  color: #000;
              }

              QMenu::item:disabled {
                  color: #999;
                  background: transparent;
              }

              QMenu::separator {
                  height: 1px;
                  background: #cccccc;
                  margin: 4px 0px;
              }
          """)
        # 显示菜单并获取选择
        action = menu.exec(self.task_list.mapToGlobal(pos))

        # 处理选择
        if action == rename_action:
            self.rename_task(task_name)
        elif action == duplicate_action:
            self.duplicate_task(task_name)
        elif action == delete_action:
            self.delete_task(task_name)

    def switch_theme(self, theme):
        if theme == "system":
            self.detect_system_theme()
        else:
            self.current_theme = theme
            self.apply_theme(theme)
            self.settings.setValue("theme", theme)

        # 更新主题菜单选中状态
        self.light_theme_action.setChecked(self.current_theme == "light")
        self.dark_theme_action.setChecked(self.current_theme == "dark")
        self.system_theme_action.setChecked(theme == "system")

        # 更新任务列表按钮样式
        for i in range(self.task_list.count()):
            item = self.task_list.item(i)
            widget = self.task_list.itemWidget(item)
            if widget:
                self.apply_button_style(widget)

    def detect_system_theme(self):
        """检测系统主题设置"""
        try:
            # 尝试检测系统是否处于暗黑模式
            # 这里只是一个示例，实际实现需要根据操作系统进行适配
            # 在Windows上可以使用注册表，在macOS上可以使用NSAppearance
            # 这里简化为使用系统设置中的值
            dark_mode = self.settings.value("systemDarkMode", False, type=bool)
            self.current_theme = "dark" if dark_mode else "light"
        except:
            self.current_theme = "light"

        self.apply_theme(self.current_theme)
        self.settings.setValue("theme", "system")

    def apply_button_style(self, widget):
        """应用按钮样式到任务项控件"""
        if self.current_theme == "light":
            widget.start_btn.setStyleSheet(self.light_button_style("start"))
            widget.stop_btn.setStyleSheet(self.light_button_style("stop"))
            widget.delete_btn.setStyleSheet(self.light_button_style("delete"))
            widget.status_label.setStyleSheet("color: #888; background: transparent;")
        else:
            widget.start_btn.setStyleSheet(self.dark_button_style("start"))
            widget.stop_btn.setStyleSheet(self.dark_button_style("stop"))
            widget.delete_btn.setStyleSheet(self.dark_button_style("delete"))
            widget.status_label.setStyleSheet("color: #aaa; background: transparent;")

    def light_button_style(self, btn_type):
        """明亮主题按钮样式"""
        base_style = """
            QPushButton {
                border-radius: 14px;
                padding: 0px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #d0d0d0;
            }
        """

        if btn_type == "start":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f5f5f5, stop:1 #e0e0e0);
            """
        elif btn_type == "stop":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f5f5f5, stop:1 #e0e0e0);
            """
        else:  # delete
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f5f5f5, stop:1 #e0e0e0);
            """

    def dark_button_style(self, btn_type):
        """暗黑主题按钮样式"""
        base_style = """
            QPushButton {
                border-radius: 14px;
                padding: 0px;
            }
            QPushButton:hover {
                background-color: #505050;
            }
            QPushButton:pressed {
                background-color: #404040;
            }
        """

        if btn_type == "start":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #505050, stop:1 #404040);
            """
        elif btn_type == "stop":
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #505050, stop:1 #404040);
            """
        else:  # delete
            return base_style + """
                background-color: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #505050, stop:1 #404040);
            """

    def apply_theme(self, theme):
        if theme == "light":
            # 明亮主题
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #f5f7fa;
                }
                QWidget {
                    background-color: #f5f7fa;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #d1d5db;
                    border-radius: 6px;
                    margin-top: 1ex;
                    background-color: white;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    top: -1px;    
                    padding: 0 5px;
                    background-color: transparent;
                    color: #333;
                }
                QListWidget, QTableWidget, QLineEdit, QComboBox, QTimeEdit, QSpinBox, QPlainTextEdit {
                    background-color: white;
                    border: 1px solid #d1d5db;
                    color: #333;
                }
                QHeaderView::section {
                    background-color: #f0f0f0;
                    color: #333;
                    border: none;
                    border-bottom: 1px solid #d1d5db;
                }
                QPushButton {
                    color: #333;
                }
                QLabel {
                    color: #333;
                }
                QCheckBox {
                    color: #333;
                    spacing: 8px;
                }
                QCheckBox::indicator {
                    width: 10px;
                    height: 10px;
                    border: 2px solid #999;
                    border-radius: 4px;
                    background: #fff;
                }
                QCheckBox::indicator:checked {
                    background: #4CAF50;
                    border: 2px solid #388E3C;
                }
                QCheckBox::indicator:hover {
                    border-color: #4CAF50;
                }
                QListWidget::item:selected {
                    background-color: #e3f2fd;
                }
            """)

            # 应用按钮样式
            self.start_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.stop_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.save_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.new_task_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
            """)

            self.clear_log_btn.setStyleSheet("""
                QPushButton {
                    background-color: #9E9E9E;
                    border: 1px solid #757575;
                    color: white;
                    padding: 2px 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #757575;
                }
            """)
        else:
            # 暗黑主题
            self.setStyleSheet("""
                QMainWindow {
                    background-color: #2d2d30;
                }
                QWidget {
                    background-color: #2d2d30;
                    color: #dcdcdc;
                }
                QGroupBox {
                    font-weight: bold;
                    border: 1px solid #3f3f46;
                    border-radius: 6px;
                    margin-top: 1ex;
                    background-color: #252526;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top center;
                    top: -1px;    
                    padding: 0 5px;
                    background-color: transparent;
                    color: #dcdcdc;
                }
                QListWidget, QTableWidget, QLineEdit, QComboBox, QTimeEdit, QSpinBox, QPlainTextEdit {
                    background-color: #252526;
                    border: 1px solid #3f3f46;
                    color: #dcdcdc;
                    selection-background-color: #3e3e40;
                }
                QHeaderView::section {
                    background-color: #3e3e40;
                    color: #dcdcdc;
                    border: none;
                    border-bottom: 1px solid #3f3f46;
                }
                QLabel {
                    color: #dcdcdc;
                }
                               QCheckBox {
                    color: #dcdcdc;
                    spacing: 8px;
                }
                QCheckBox::indicator {
                    width: 10px;
                    height: 10px;
                    border: 2px solid #666;
                    border-radius: 4px;
                    background: #2d2d30;
                }
                QCheckBox::indicator:checked {
                    background: #4CAF50;
                    border: 2px solid #388E3C;
                }
                QCheckBox::indicator:hover {
                    border-color: #4CAF50;
                }
                QListWidget::item:selected {
                    background-color: #3e3e40;
                }
            """)

            # 应用按钮样式
            self.start_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.stop_current_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.save_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
                QPushButton:disabled {
                    background-color: #81C784;
                }
            """)

            self.new_task_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    border: 1px solid #388E3C;
                    color: white;
                    padding: 5px;
                    border-radius: 4px;
                }
                QPushButton:hover {
                    background-color: #388E3C;
                }
            """)

            self.clear_log_btn.setStyleSheet("""
                QPushButton {
                    background-color: #9E9E9E;
                    border: 1px solid #757575;
                    color: white;
                    padding: 2px 5px;
                    border-radius: 3px;
                }
                QPushButton:hover {
                    background-color: #757575;
                }
            """)

    def load_settings(self):
        # 加载主题设置
        self.current_theme = self.settings.value("theme", "light")

    def save_settings(self):
        # 保存分割器位置
        self.settings.setValue("splitterSizes", self.splitter.sizes())

    def closeEvent(self, event):
        self.save_settings()
        # 停止所有定时器
        for timer in self.scheduled_timers.values():
            timer.stop()
        event.accept()

    def add_step(self):
        # 检查是否有任务
        if self.task_list.count() == 0:
            QMessageBox.information(self, "提示", "请先新建任务，再添加步骤！")
            return
        
        # 检查是否有当前选中的任务
        if not self.current_task or self.current_task not in self.tasks:
            QMessageBox.information(self, "提示", "请先选中一个任务，再添加步骤！")
            return
            
        dialog = StepConfigDialog(parent=self)
        if dialog.exec() == QDialog.Accepted:
            step_data = dialog.get_step_data()
            self.add_step_to_table(step_data)

            # 添加到当前任务配置
            if self.current_task and self.current_task in self.tasks:
                self.tasks[self.current_task]["steps"].append(step_data)

    def copy_step(self):
        """复制当前选中的步骤"""
        selected_row = self.steps_table.currentRow()
        if selected_row < 0:
            QMessageBox.information(self, "提示", "请先选中一条步骤再复制。")
            return

        # 取出原步骤数据
        src_step = self.tasks[self.current_task]["steps"][selected_row]
        # 深拷贝，避免后续修改互相影响
        new_step = deepcopy(src_step)

        # 直接追加到表格和任务配置
        self.add_step_to_table(new_step)
        self.tasks[self.current_task]["steps"].append(new_step)
    def edit_step(self):
        selected_row = self.steps_table.currentRow()
        if selected_row < 0:
            return

        # 获取当前步骤数据
        step_data = self.tasks[self.current_task]["steps"][selected_row]
        dialog = StepConfigDialog(step_data,parent=self)
        if dialog.exec() == QDialog.Accepted:
            new_step_data = dialog.get_step_data()
            # 更新表格
            use_color = self.label_color_checkbox.isChecked() if hasattr(self, 'label_color_checkbox') else True
            type_widget = StepTableHelper.type_widget(new_step_data["type"], use_color)
            self.steps_table.setCellWidget(selected_row, 0, type_widget)
            # self.steps_table.setItem(selected_row, 0, QTableWidgetItem(new_step_data["type"]))
            w = StepTableHelper.widget_of(new_step_data,use_color)
            self.steps_table.setCellWidget(selected_row, 1, w)
            self.steps_table.setRowHeight(selected_row, max(StepTableHelper.IMG_HEIGHT + 4, 24))
            self.steps_table.verticalHeader().setDefaultSectionSize(
                StepTableHelper.FIXED_ROW_HEIGHT
            )
            self.steps_table.horizontalHeader().setStretchLastSection(True)
            # 格式化参数显示
            params_text = ""
            params = new_step_data["params"]
            if new_step_data["type"] == "鼠标点击":
                use_image = params.get('use_image', True)
                use_coordinates = params.get('use_coordinates', False)

                if use_image:
                    img_path = params.get('image_path', '')
                    click_type = params.get('click_type', '左键单击')
                    scan_direction = params.get('scan_direction', '默认')
                    offset_x = params.get('offset_x', 0)
                    offset_y = params.get('offset_y', 0)

                    img_name = os.path.basename(img_path) if img_path else "未设置"
                    params_text = f"图片: {img_name}, 点击: {click_type}, 方向: {scan_direction}"
                    if offset_x != 0 or offset_y != 0:
                        params_text += f", 偏移: ({offset_x}, {offset_y})"

                elif use_coordinates:
                    x_coord = params.get('x_coordinate', 0)
                    y_coord = params.get('y_coordinate', 0)
                    click_type = params.get('click_type', '左键单击')
                    offset_x = params.get('offset_x', 0)
                    offset_y = params.get('offset_y', 0)

                    params_text = f"坐标: ({x_coord}, {y_coord}), 点击: {click_type}"
                    if offset_x != 0 or offset_y != 0:
                        params_text += f", 偏移: ({offset_x}, {offset_y})"

                else:
                    params_text = "未启用图片或坐标模式"
            elif new_step_data["type"] == "文本输入":
                # 优先显示纯文本
                txt = params.get("text", "")
                if txt:
                    params_text = f"文本: {txt}"
                else:
                    # Excel 模式
                    mode = params.get("mode", "顺序")
                    path = os.path.basename(params.get("excel_path", ""))
                    sheet = params.get("sheet", "0")
                    col = params.get("col", 0)
                    params_text = f"Excel({mode}) {path}|{sheet}|列{col}"
            elif new_step_data["type"] == "等待":
                params_text = f"等待: {new_step_data['params'].get('seconds', 0)}秒"
            elif new_step_data["type"] == "截图":
                params_text = f"保存到: {new_step_data['params'].get('save_path', '')}"
            elif new_step_data["type"] == "拖拽":
                use_image = new_step_data['params'].get('use_image', True)
                if use_image:
                    img_path = new_step_data['params'].get('image_path', '')
                    if img_path:
                        img_name = os.path.basename(img_path)
                        dx = new_step_data['params'].get('drag_x', 0)
                        dy = new_step_data['params'].get('drag_y', 0)
                        params_text = f"图片: {img_name} (+{dx},+{dy})"
                    else:
                        params_text = "图片: 未设置"
                else:
                    start_x = new_step_data['params'].get('start_x', 0)
                    start_y = new_step_data['params'].get('start_y', 0)
                    end_x = new_step_data['params'].get('end_x', 0)
                    end_y = new_step_data['params'].get('end_y', 0)
                    params_text = f"从({start_x},{start_y})到({end_x},{end_y})"
            elif new_step_data["type"] == "自动回复":
                    provider = new_step_data['params'].get("provider", "kimi")
                    system_prompt = new_step_data['params'].get("system_prompt", "")
                    use_history = new_step_data['params'].get("use_history", True)
                    stream = new_step_data['params'].get("stream", False)

                    # 构造参数描述文本
                    params_text = (
                        f"AI 提供商: {provider}, "
                        f"系统提示词: {'已设置' if system_prompt else '未设置'}, "
                        f"使用历史记录: {'是' if use_history else '否'}, "
                        f"流式输出: {'是' if stream else '否'}"
                    )
            self.steps_table.setItem(selected_row, 2, QTableWidgetItem(params_text))
            # self.steps_table.setItem(selected_row, 3, QTableWidgetItem(str(new_step_data.get("delay", 0))))

            self.tasks[self.current_task]["steps"][selected_row] = new_step_data



    def remove_step(self):
        selected_row = self.steps_table.currentRow()
        if selected_row >= 0:
            self.steps_table.removeRow(selected_row)

            # 从任务配置中移除
            if self.current_task and self.current_task in self.tasks:
                self.tasks[self.current_task]["steps"].pop(selected_row)

    def move_step_up(self):
        selected_row = self.steps_table.currentRow()
        if selected_row > 0:
            # 移动表格行
            self.steps_table.insertRow(selected_row - 1)
            for col in range(self.steps_table.columnCount()):
                # 移动 QTableWidgetItem
                self.steps_table.setItem(selected_row - 1, col, self.steps_table.takeItem(selected_row + 1, col))
                # 移动 cellWidget
                widget = self.steps_table.cellWidget(selected_row + 1, col)
                if widget:
                    self.steps_table.setCellWidget(selected_row - 1, col, widget)
            self.steps_table.removeRow(selected_row + 1)
            self.steps_table.setCurrentCell(selected_row - 1, 0)

            # 移动任务配置中的步骤
            if self.current_task and self.current_task in self.tasks:
                steps = self.tasks[self.current_task]["steps"]
                steps.insert(selected_row - 1, steps.pop(selected_row))

    def move_step_down(self):
        selected_row = self.steps_table.currentRow()
        if selected_row >= 0 and selected_row < self.steps_table.rowCount() - 1:
            # 移动表格行
            self.steps_table.insertRow(selected_row + 2)
            for col in range(self.steps_table.columnCount()):
                # 移动 QTableWidgetItem
                self.steps_table.setItem(selected_row + 2, col, self.steps_table.takeItem(selected_row, col))
                # 移动 cellWidget
                widget = self.steps_table.cellWidget(selected_row, col)
                if widget:
                    self.steps_table.setCellWidget(selected_row + 2, col, widget)
            self.steps_table.removeRow(selected_row)
            self.steps_table.setCurrentCell(selected_row + 1, 0)

            # 移动任务配置中的步骤
            if self.current_task and self.current_task in self.tasks:
                steps = self.tasks[self.current_task]["steps"]
                steps.insert(selected_row + 1, steps.pop(selected_row))

    def save_task_config(self):
        if not self.current_task:
            return

        # 更新任务名称
        new_name = self.task_name.text().strip()
        if new_name and new_name != self.current_task:
            # 更新任务列表
            for i in range(self.task_list.count()):
                item = self.task_list.item(i)
                widget = self.task_list.itemWidget(item)
                if widget and widget.task_name == self.current_task:
                    widget.task_name = new_name
                    widget.name_label.setText(new_name)

                    # 更新任务配置
                    task_config = self.tasks.pop(self.current_task)
                    task_config["name"] = new_name
                    self.tasks[new_name] = task_config
                    self.current_task = new_name
                    break

        self.export_config_default()
        # QMessageBox.information(self, "保存成功", "任务配置已保存")

    def export_config(self):
        if not self.current_task:
            return
        if self.current_task in self.tasks:
            self.tasks[self.current_task]["schedule"] = {
                "enable": self.schedule_enable.currentText(),
                "time": self.schedule_time.time().toString("HH:mm:ss"),
                "interval": self.repeat_interval.value(),
                "repeat": self.repeat_count.currentText()
            }
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出配置", "", "JSON文件 (*.json)"
        )
        if file_path:
            if not file_path.lower().endswith('.json'):
                file_path += '.json'

            if self.current_task in self.tasks:
                with open(file_path, 'w') as f:
                    json.dump(self.tasks[self.current_task], f, indent=4)
                QMessageBox.information(self, "导出成功", "任务配置已导出")

    def export_config_default(self):
        if not self.current_task:
            return
        if self.current_task in self.tasks:
            self.tasks[self.current_task]["schedule"] = {
                "enable": self.schedule_enable.currentText(),
                "time": self.schedule_time.time().toString("HH:mm:ss"),
                "interval": self.repeat_interval.value(),
                "repeat": self.repeat_count.currentText()
            }

            # 创建config目录（如果不存在）
            config_dir = os.path.join(os.getcwd(), "config")
            os.makedirs(config_dir, exist_ok=True)

            # 生成文件路径
            file_name = f"{self.current_task}.json"
            file_path = os.path.join(config_dir, file_name)

            # 保存配置文件
            if self.current_task in self.tasks:
                with open(file_path, 'w') as f:
                    json.dump(self.tasks[self.current_task], f, indent=4)
                QMessageBox.information(self, "导出成功", f"任务配置已导出到: {file_path}")

    def import_config(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "导入配置", "", "JSON文件 (*.json)"
        )
        if file_path:
            try:
                with open(file_path, 'r') as f:
                    task_config = json.load(f)

                task_name = task_config.get("name", "导入的任务")
                self.add_task(task_name)
                self.tasks[task_name] = task_config

                # 选中新导入的任务
                for i in range(self.task_list.count()):
                    item = self.task_list.item(i)
                    widget = self.task_list.itemWidget(item)
                    if widget and widget.task_name == task_name:
                        self.task_list.setCurrentItem(item)
                        break
                self.load_task_config(task_name)

                QMessageBox.information(self, "导入成功", "任务配置已导入")
            except Exception as e:
                QMessageBox.critical(self, "导入失败", f"导入配置时出错: {str(e)}")

    def apply_schedule(self):
        """应用定时设置"""
        if not self.current_task:
            return

        task_name = self.current_task

        # 如果任务已有定时器，先停止
        if task_name in self.scheduled_timers:
            self.scheduled_timers[task_name].stop()
            del self.scheduled_timers[task_name]

        # 获取定时设置
        schedule_type = self.schedule_enable.currentText()
        if schedule_type == "立即执行":
            # 不需要定时器
            return

        # 定时执行
        schedule_time = self.schedule_time.time()
        interval_minutes = self.repeat_interval.value()
        repeat_count = self.repeat_count.currentText()

        # 计算第一次执行的时间
        now = QTime.currentTime()
        first_run = QTime(schedule_time.hour(), schedule_time.minute(), schedule_time.second())

        # 如果当前时间已超过设定时间，则明天执行
        if first_run < now:
            first_run = first_run.addSecs(24 * 3600)  # 加一天

        # 计算延迟时间（毫秒）
        delay_ms = now.msecsTo(first_run)

        # 创建首次执行的定时器
        initial_timer = QTimer(self)
        initial_timer.setSingleShot(True)  # 只执行一次

        def run_initial_task():
            # 执行任务
            self.start_current_task()

            # 如果需要重复执行，设置重复定时器
            if repeat_count == "无限":
                repeat_timer = QTimer(self)
                repeat_timer.timeout.connect(self.start_current_task)
                repeat_timer.setInterval(interval_minutes * 60 * 1000)  # 转换为毫秒
                repeat_timer.start()
                # 保存重复定时器引用
                self.scheduled_timers[task_name] = repeat_timer
            elif repeat_count != "1":
                try:
                    total_count = int(repeat_count)
                    current_count = 1  # 第一次已经执行

                    if current_count < total_count:
                        repeat_timer = QTimer(self)

                        def run_repeat_task():
                            nonlocal current_count
                            self.start_current_task()
                            current_count += 1
                            if current_count >= total_count:
                                repeat_timer.stop()
                                if task_name in self.scheduled_timers:
                                    del self.scheduled_timers[task_name]

                        repeat_timer.timeout.connect(run_repeat_task)
                        repeat_timer.setInterval(interval_minutes * 60 * 1000)
                        repeat_timer.start()
                        # 保存重复定时器引用
                        self.scheduled_timers[task_name] = repeat_timer
                except ValueError:
                    pass  # 无效的重复次数

        initial_timer.timeout.connect(run_initial_task)
        initial_timer.start(delay_ms)

        # 保存首次执行定时器引用
        self.scheduled_timers[task_name] = initial_timer

        # 显示提示信息
        self.log_text.appendPlainText(
            f"[{time.strftime('%H:%M:%S')}] 已设置定时任务: {task_name} 将在 {first_run.toString('HH:mm:ss')} 执行")
        QMessageBox.information(self, "定时成功",  f"[{time.strftime('%H:%M:%S')}] 已设置定时任务: {task_name} 将在 {first_run.toString('HH:mm:ss')} 执行\n请保持桌面处于从不熄屏状态")
    def run_scheduled_task(self, task_name, timer, count):
        """执行定时任务（带计数）"""
        if count <= 0:
            timer.stop()
            if task_name in self.scheduled_timers:
                del self.scheduled_timers[task_name]
            return

        # 执行任务
        self.start_current_task()

        # 减少计数
        if count > 1:
            # 设置下一次执行
            QTimer.singleShot(0, lambda: self.run_scheduled_task(task_name, timer, count - 1))
        else:
            timer.stop()
            if task_name in self.scheduled_timers:
                del self.scheduled_timers[task_name]


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutomationUI()
    window.setWindowIcon(ATIcon.icon())
    window.show()
    sys.exit(app.exec())
