# -*- coding: utf-8 -*-
import os
import sys
import re
import threading
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, QMessageBox,
    QFrame
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont
import pdfplumber

# ========== 页码解析 ==========
def parse_page_range(page_expr: str, total_pages: int):
    selected_pages = set()
    if not page_expr.strip():
        return list(range(1, total_pages + 1))
    parts = page_expr.split(",")
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            try:
                start_str, end_str = part.split("-", 1)
                start = int(start_str.strip())
                end = int(end_str.strip())
            except ValueError:
                raise ValueError(f"页码格式错误：{part}（例：1-10）")
            start = max(1, start)
            end = min(total_pages, end)
            if start > end:
                raise ValueError(f"起始页不能大于结束页：{part}")
            for p in range(start, end + 1):
                selected_pages.add(p)
        else:
            try:
                page = int(part)
            except ValueError:
                raise ValueError(f"页码不是数字：{part}")
            if 1 <= page <= total_pages:
                selected_pages.add(page)
    return sorted(list(selected_pages))

# ========== 统计线程（避免界面卡顿） ==========
class CountThread(QThread):
    progress_signal = pyqtSignal(str)
    result_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)

    def __init__(self, pdf_path, page_expr):
        super().__init__()
        self.pdf_path = pdf_path
        self.page_expr = page_expr

    def run(self):
        try:
            self.progress_signal.emit("正在读取PDF文件...")
            # 获取总页数
            with pdfplumber.open(self.pdf_path) as pdf:
                total_pages = len(pdf.pages)
            self.progress_signal.emit(f"解析页码范围...（总页数：{total_pages}）")
            # 解析页码
            selected_pages = parse_page_range(self.page_expr, total_pages)
            # 提取文本（优化版：禁用布局分析提速）
            self.progress_signal.emit("提取PDF文本...")
            text = ""
            with pdfplumber.open(self.pdf_path) as pdf:
                all_pages = pdf.pages
                for idx in [p-1 for p in selected_pages]:
                    page = all_pages[idx]
                    # 只提取纯文本，禁用布局分析（提速30%+）
                    page_text = page.extract_text(layout=False) or ""
                    text += page_text

            if not text.strip():
                self.error_signal.emit("未提取到文本！可能是图片版PDF，无法统计。")
                return

            # 统计（Word 标准）
            self.progress_signal.emit("按Word规则统计...")
            clean_no_space = text.replace(" ", "").replace("\n", "").replace("\r", "").replace("\t", "")
            clean_with_space = text.replace("\n", " ").replace("\r", " ").replace("\t", " ")
            file_size = os.path.getsize(self.pdf_path) / 1024 / 1024

            # Word 字数
            cn_full_pat = r'[\u4e00-\u9fff\u3000-\u303f\uff00-\uffef]'
            cn_full = re.findall(cn_full_pat, text)
            en_words = re.findall(r'[a-zA-Z]+(?:[\'-][a-zA-Z]+)*', clean_with_space)
            num_words = re.findall(r'\d+(?:\.\d+)?', clean_with_space)
            total_words = len(cn_full) + len(en_words) + len(num_words)

            # 字符数
            chars_no_space = len(clean_no_space)
            space_count = len([c for c in text if c in " \n\r\t"])
            chars_with_space = chars_no_space + space_count

            # 行数
            raw_lines = len(text.splitlines())
            non_empty_lines = len([line for line in text.splitlines() if line.strip()])

            # 结果封装
            result = {
                "file_name": os.path.basename(self.pdf_path),
                "file_path": self.pdf_path,
                "file_size": round(file_size, 2),
                "page_count": total_pages,
                "selected_pages": selected_pages,
                "selected_page_count": len(selected_pages),
                "total_words": total_words,
                "chars_no_space": chars_no_space,
                "chars_with_space": chars_with_space,
                "cn_full_count": len(cn_full),
                "en_word_count": len(en_words),
                "num_word_count": len(num_words),
                "space_count": space_count,
                "raw_lines": raw_lines,
                "non_empty_lines": non_empty_lines,
                "count_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.result_signal.emit(result)
        except Exception as e:
            self.error_signal.emit(f"统计失败：{str(e)}")

# ========== 主窗口（现代UI+拖拽功能） ==========
class PDFWordCountWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF字数统计工具")
        
        # ========== 修正图标路径 ==========
        import os
        from PyQt6.QtGui import QIcon
        # 读取ICO图标（Windows兼容）
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        self.setMinimumSize(920, 780)
        self.resize(920, 780)
        self.setStyleSheet("""
            QMainWindow { background-color: #fafbfc; }
            QFrame { background-color: white; border-radius: 8px; }
            QLabel { font-family: "微软雅黑"; font-size: 11pt; color: #2c3e50; }
            QLabel#title { font-size: 18pt; font-weight: bold; color: #2c3e50; }
            QLabel#subtitle { font-size: 9pt; color: #7f8c8d; }
            QLabel#hint { font-size: 9pt; color: #95a5a6; }
            QLineEdit { 
                font-family: "微软雅黑"; font-size: 10pt; 
                border: 1px solid #e0e6ed; border-radius: 4px;
                padding: 6px; background-color: white;
            }
            QLineEdit:focus { border-color: #2d8bfd; }
            QPushButton {
                font-family: "微软雅黑"; font-size: 10pt;
                border-radius: 4px; padding: 8px 16px;
            }
            QPushButton#primary {
                background-color: #2d8bfd; color: white;
                font-weight: bold;
            }
            QPushButton#primary:hover { background-color: #1b7cef; }
            QPushButton#primary:pressed { background-color: #0a68df; }
            QPushButton#normal {
                background-color: #f5f7fa; color: #2c3e50;
                border: 1px solid #e0e6ed;
            }
            QPushButton#normal:hover { background-color: #e8eff8; }
            QTextEdit {
                font-family: "微软雅黑"; font-size: 10pt;
                border: none; background-color: white;
                padding: 10px;
            }
            QScrollBar:vertical {
                width: 8px; background-color: #f5f7fa;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background-color: #cbd5e1; border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover { background-color: #94a3b8; }
        """)

        # 中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(36, 22, 36, 22)
        main_layout.setSpacing(12)

        # 标题
        title_layout = QVBoxLayout()
        title_label = QLabel("📄 PDF字数统计工具")
        title_label.setObjectName("title")
        subtitle_label = QLabel("统计规则参考 Microsoft Word  | 适配 Windows 10/11")
        subtitle_label.setObjectName("subtitle")
        title_layout.addWidget(title_label, alignment=Qt.AlignmentFlag.AlignCenter)
        title_layout.addWidget(subtitle_label, alignment=Qt.AlignmentFlag.AlignCenter)
        main_layout.addLayout(title_layout)

        # 文件选择卡片（支持拖拽）
        file_frame = QFrame()
        file_layout = QVBoxLayout(file_frame)
        file_layout.setContentsMargins(24, 18, 24, 18)
        file_layout.setSpacing(10)

        # 文件路径行
        path_layout = QHBoxLayout()
        path_label = QLabel("PDF文件路径：")
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("点击“浏览文件”或直接拖拽PDF文件到此处...")
        self.path_edit.setAcceptDrops(True)  # 启用拖拽
        self.path_edit.dragEnterEvent = self.on_drag_enter  # 拖拽进入
        self.path_edit.dropEvent = self.on_drop  # 释放文件
        browse_btn = QPushButton("浏览文件")
        browse_btn.setObjectName("normal")
        browse_btn.clicked.connect(self.select_file)
        path_layout.addWidget(path_label)
        path_layout.addWidget(self.path_edit)
        path_layout.addWidget(browse_btn)
        file_layout.addLayout(path_layout)

        main_layout.addWidget(file_frame)

        # 页码选择卡片
        page_frame = QFrame()
        page_layout = QVBoxLayout(page_frame)
        page_layout.setContentsMargins(24, 18, 24, 18)
        page_layout.setSpacing(10)

        page_h_layout = QHBoxLayout()
        page_label = QLabel("指定页码：")
        self.page_edit = QLineEdit()
        self.page_edit.setPlaceholderText("例：1-10,12,15-20（留空统计全部）")
        hint_label = QLabel("格式：1-100 或 1,2-33 或 1-3,6,8,10-66")
        hint_label.setObjectName("hint")
        page_h_layout.addWidget(page_label)
        page_h_layout.addWidget(self.page_edit)
        page_h_layout.addWidget(hint_label)
        page_layout.addLayout(page_h_layout)

        main_layout.addWidget(page_frame)

        # 按钮行
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        self.count_btn = QPushButton("开始统计")
        self.count_btn.setObjectName("primary")
        self.count_btn.clicked.connect(self.start_count)
        clear_btn = QPushButton("清空结果")
        clear_btn.setObjectName("normal")
        clear_btn.clicked.connect(self.clear_result)
        save_btn = QPushButton("保存结果")
        save_btn.setObjectName("normal")
        save_btn.clicked.connect(self.save_result)
        btn_layout.addWidget(self.count_btn)
        btn_layout.addWidget(clear_btn)
        btn_layout.addWidget(save_btn)
        btn_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addLayout(btn_layout)

        # 结果卡片
        result_frame = QFrame()
        result_layout = QVBoxLayout(result_frame)
        result_layout.setContentsMargins(24, 20, 24, 20)
        result_layout.setSpacing(10)

        result_title = QLabel("📊 统计结果（Word 官方标准）")
        result_title.setStyleSheet("font-size: 12pt; font-weight: bold;")
        self.result_edit = QTextEdit()
        self.result_edit.setReadOnly(True)
        self.result_edit.setMinimumHeight(300)

        result_layout.addWidget(result_title)
        result_layout.addWidget(self.result_edit)
        main_layout.addWidget(result_frame, stretch=1)

        # 统计线程
        self.count_thread = None

    # 拖拽功能（Windows 10/11 原生支持）
    def on_drag_enter(self, event):
        """拖拽进入时接受文件"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.path_edit.setStyleSheet("border-color: #2d8bfd;")

    def on_drop(self, event):
        """释放文件时获取路径"""
        self.path_edit.setStyleSheet("border: 1px solid #e0e6ed;")
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(".pdf"):
                self.path_edit.setText(file_path)

    # 选择文件
    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择PDF文件", "", "PDF文件 (*.pdf);;所有文件 (*.*)"
        )
        if file_path:
            self.path_edit.setText(file_path)

    # 清空结果（保留文件和页码）
    def clear_result(self):
        self.result_edit.clear()

    # 保存结果
    def save_result(self):
        content = self.result_edit.toPlainText().strip()
        if not content:
            QMessageBox.warning(self, "提示", "暂无统计结果可保存！")
            return
        
        # 构造默认保存路径和文件名（基于原PDF文件）
        pdf_path = self.path_edit.text().strip()
        if pdf_path and os.path.exists(pdf_path):
            # 提取PDF文件的目录和主名（去掉扩展名）
            pdf_dir = os.path.dirname(pdf_path)
            pdf_base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            # 拼接默认文件名：原PDF名_字数统计.txt
            default_save_path = os.path.join(pdf_dir, f"{pdf_base_name}_字数统计.txt")
        else:
            # 兜底：使用原有的时间戳命名
            default_save_path = f"PDF字数统计_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        save_path, _ = QFileDialog.getSaveFileName(
            self, "保存统计结果",
            default_save_path,  # 使用构造的默认路径
            "文本文件 (*.txt);;所有文件 (*.*)"
        )
        if save_path:
            try:
                with open(save_path, "w", encoding="utf-8") as f:
                    f.write(content)
                QMessageBox.information(self, "成功", f"结果已保存到：\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存失败：{str(e)}")

    # 开始统计
    def start_count(self):
        pdf_path = self.path_edit.text().strip()
        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "提示", "请选择有效的PDF文件！")
            return
        if not pdf_path.lower().endswith(".pdf"):
            QMessageBox.warning(self, "提示", "请选择PDF格式的文件！")
            return

        # 禁用按钮
        self.count_btn.setEnabled(False)
        self.count_btn.setText("统计中...")
        self.result_edit.clear()

        # 启动统计线程
        self.count_thread = CountThread(pdf_path, self.page_edit.text().strip())
        self.count_thread.progress_signal.connect(self.update_progress)
        self.count_thread.result_signal.connect(self.show_result)
        self.count_thread.error_signal.connect(self.show_error)
        self.count_thread.finished.connect(self.count_finished)
        self.count_thread.start()

    # 更新进度
    def update_progress(self, msg):
        self.result_edit.append(f"🔄 {msg}")

    # 显示结果
    def show_result(self, result):
        report = f"""
{'='*70}
📊 PDF字数统计报告
{'='*70}
【文件信息】
文件名：{result['file_name']}
文件大小：{result['file_size']} MB
PDF总页数：{result['page_count']} 页
本次统计：{result['selected_page_count']} 页
统计页码：{result['selected_pages']}

【核心统计结果】
✅ 字数：{result['total_words']:,}
✅ 字符数（不含空格）：{result['chars_no_space']:,}
✅ 字符数（含空格）：{result['chars_with_space']:,}

【明细构成】
中文+全角标点：{result['cn_full_count']:,}
英文单词：{result['en_word_count']:,}
数字串：{result['num_word_count']:,}
空格/换行符：{result['space_count']:,}
总行数：{result['raw_lines']:,}
非空行数：{result['non_empty_lines']:,}

统计时间：{result['count_time']}
{'='*70}
💡 结果可直接用于论文/报告字数核对
        """
        self.result_edit.setText(report)

    # 显示错误
    def show_error(self, msg):
        self.result_edit.setText(f"❌ {msg}")

    # 统计完成恢复按钮
    def count_finished(self):
        self.count_btn.setEnabled(True)
        self.count_btn.setText("开始统计")

# ========== 主函数 ==========
if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 完全移除高DPI手动配置（Qt6自动适配，避免属性报错）
    window = PDFWordCountWindow()
    window.show()
    sys.exit(app.exec())