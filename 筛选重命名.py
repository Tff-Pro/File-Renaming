import os
import sys
import glob
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QPushButton, QLabel, QLineEdit, QComboBox, QListWidget,
                             QFileDialog, QMessageBox, QGroupBox, QCheckBox, QSpinBox,
                             QProgressBar, QTextEdit, QSplitter)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

class FileRenamerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.folder_path = ""
        
    def initUI(self):
        self.setWindowTitle("文件重命名工具")
        self.setGeometry(100, 100, 900, 700)
        
        # 中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 创建分割器
        splitter = QSplitter(Qt.Vertical)
        
        # 上部 - 控制面板
        control_group = QGroupBox("重命名设置")
        control_layout = QVBoxLayout()
        
        # 文件夹选择
        folder_layout = QHBoxLayout()
        self.folder_label = QLabel("未选择文件夹")
        self.folder_label.setStyleSheet("QLabel { background-color: #f0f0f0; padding: 5px; }")
        folder_btn = QPushButton("选择文件夹")
        folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.folder_label, 4)
        folder_layout.addWidget(folder_btn, 1)
        control_layout.addLayout(folder_layout)
        
        # 文件类型选择
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("文件类型:"))
        self.file_type_combo = QComboBox()
        self.file_type_combo.addItems(["所有文件", "PDF文件", "图片文件", "文本文件", "视频文件", "音频文件", "自定义"])
        self.file_type_combo.currentTextChanged.connect(self.update_file_list)
        type_layout.addWidget(self.file_type_combo)
        
        self.custom_ext_input = QLineEdit()
        self.custom_ext_input.setPlaceholderText("输入扩展名，如: txt,docx (逗号分隔)")
        self.custom_ext_input.setVisible(False)
        type_layout.addWidget(self.custom_ext_input)
        
        control_layout.addLayout(type_layout)
        
        # 排序方式
        sort_layout = QHBoxLayout()
        sort_layout.addWidget(QLabel("排序方式:"))
        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["修改时间(旧→新)", "修改时间(新→旧)", "文件名(A→Z)", "文件名(Z→A)", "文件大小(小→大)", "文件大小(大→小)"])
        sort_layout.addWidget(self.sort_combo)
        control_layout.addLayout(sort_layout)
        
        # 重命名规则
        rename_layout = QHBoxLayout()
        rename_layout.addWidget(QLabel("重命名规则:"))
        self.rename_combo = QComboBox()
        self.rename_combo.addItems(["序号+原文件名", "完全重命名", "日期+序号", "自定义格式"])
        self.rename_combo.currentTextChanged.connect(self.toggle_custom_format)
        rename_layout.addWidget(self.rename_combo)
        
        self.custom_format_input = QLineEdit()
        self.custom_format_input.setPlaceholderText("使用 {index} {name} {ext} {date} 等变量")
        self.custom_format_input.setVisible(False)
        rename_layout.addWidget(self.custom_format_input)
        control_layout.addLayout(rename_layout)
        
        # 序号设置
        index_layout = QHBoxLayout()
        index_layout.addWidget(QLabel("起始序号:"))
        self.start_index_spin = QSpinBox()
        self.start_index_spin.setRange(1, 9999)
        self.start_index_spin.setValue(1)
        index_layout.addWidget(self.start_index_spin)
        
        index_layout.addWidget(QLabel("序号位数:"))
        self.digits_spin = QSpinBox()
        self.digits_spin.setRange(1, 6)
        self.digits_spin.setValue(3)
        index_layout.addWidget(self.digits_spin)
        control_layout.addLayout(index_layout)
        
        # 预览和操作按钮
        btn_layout = QHBoxLayout()
        self.preview_btn = QPushButton("预览重命名")
        self.preview_btn.clicked.connect(self.preview_rename)
        self.rename_btn = QPushButton("执行重命名")
        self.rename_btn.clicked.connect(self.execute_rename)
        self.rename_btn.setEnabled(False)
        
        btn_layout.addWidget(self.preview_btn)
        btn_layout.addWidget(self.rename_btn)
        control_layout.addLayout(btn_layout)
        
        control_group.setLayout(control_layout)
        
        # 下部 - 文件列表和日志
        bottom_group = QGroupBox("文件列表和操作日志")
        bottom_layout = QVBoxLayout()
        
        # 文件列表
        self.file_list = QListWidget()
        bottom_layout.addWidget(QLabel("文件列表:"))
        bottom_layout.addWidget(self.file_list)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        bottom_layout.addWidget(self.progress_bar)
        
        # 日志区域
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        bottom_layout.addWidget(QLabel("操作日志:"))
        bottom_layout.addWidget(self.log_text)
        
        bottom_group.setLayout(bottom_layout)
        
        # 添加到分割器
        splitter.addWidget(control_group)
        splitter.addWidget(bottom_group)
        splitter.setSizes([300, 400])
        
        main_layout.addWidget(splitter)
        
        # 连接信号
        self.file_type_combo.currentTextChanged.connect(self.toggle_custom_ext)
        self.custom_ext_input.textChanged.connect(self.update_file_list)
        self.sort_combo.currentTextChanged.connect(self.update_file_list)
        
    def toggle_custom_ext(self):
        """切换自定义扩展名输入框的显示"""
        if self.file_type_combo.currentText() == "自定义":
            self.custom_ext_input.setVisible(True)
        else:
            self.custom_ext_input.setVisible(False)
            
    def toggle_custom_format(self):
        """切换自定义格式输入框的显示"""
        if self.rename_combo.currentText() == "自定义格式":
            self.custom_format_input.setVisible(True)
        else:
            self.custom_format_input.setVisible(False)
            
    def select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.folder_path = folder
            self.folder_label.setText(folder)
            self.update_file_list()
            
    def get_file_extensions(self):
        """获取选择的文件扩展名"""
        file_type = self.file_type_combo.currentText()
        if file_type == "所有文件":
            return ["*"]
        elif file_type == "PDF文件":
            return ["pdf"]
        elif file_type == "图片文件":
            return ["jpg", "jpeg", "png", "gif", "bmp", "tiff"]
        elif file_type == "文本文件":
            return ["txt", "doc", "docx", "md", "rtf"]
        elif file_type == "视频文件":
            return ["mp4", "avi", "mov", "mkv", "wmv"]
        elif file_type == "音频文件":
            return ["mp3", "wav", "flac", "aac", "ogg"]
        elif file_type == "自定义":
            custom_ext = self.custom_ext_input.text().strip()
            if custom_ext:
                return [ext.strip() for ext in custom_ext.split(",")]
        return ["*"]
            
    def update_file_list(self):
        """更新文件列表"""
        if not self.folder_path:
            return
            
        self.file_list.clear()
        extensions = self.get_file_extensions()
        
        # 获取所有匹配的文件
        all_files = []
        for ext in extensions:
            if ext == "*":
                pattern = "*"
            else:
                pattern = f"*.{ext}"
            files = glob.glob(os.path.join(self.folder_path, pattern))
            all_files.extend([f for f in files if os.path.isfile(f)])
            
        # 去重
        all_files = list(set(all_files))
        
        # 排序
        sort_method = self.sort_combo.currentText()
        if sort_method == "修改时间(旧→新)":
            all_files.sort(key=os.path.getmtime)
        elif sort_method == "修改时间(新→旧)":
            all_files.sort(key=os.path.getmtime, reverse=True)
        elif sort_method == "文件名(A→Z)":
            all_files.sort(key=lambda x: os.path.basename(x).lower())
        elif sort_method == "文件名(Z→A)":
            all_files.sort(key=lambda x: os.path.basename(x).lower(), reverse=True)
        elif sort_method == "文件大小(小→大)":
            all_files.sort(key=os.path.getsize)
        elif sort_method == "文件大小(大→小)":
            all_files.sort(key=os.path.getsize, reverse=True)
            
        # 显示文件列表
        for file_path in all_files:
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            item_text = f"{file_name} ({file_size//1024}KB, {mod_time.strftime('%Y-%m-%d %H:%M')})"
            self.file_list.addItem(item_text)
            
        self.log_text.append(f"找到 {len(all_files)} 个文件")
        
    def generate_new_name(self, old_name, index):
        """生成新文件名"""
        name_without_ext, ext = os.path.splitext(old_name)
        ext = ext.lower()
        
        index_str = str(index).zfill(self.digits_spin.value())
        current_date = datetime.now().strftime("%Y%m%d")
        
        rule = self.rename_combo.currentText()
        
        if rule == "序号+原文件名":
            return f"{index_str}_{old_name}"
        elif rule == "完全重命名":
            return f"{index_str}{ext}"
        elif rule == "日期+序号":
            return f"{current_date}_{index_str}{ext}"
        elif rule == "自定义格式":
            format_str = self.custom_format_input.text()
            if format_str:
                return format_str.format(
                    index=index_str,
                    name=name_without_ext,
                    ext=ext[1:],  # 去掉点
                    date=current_date,
                    original=old_name
                )
            return old_name
        return old_name
        
    def preview_rename(self):
        """预览重命名结果"""
        if not self.folder_path:
            QMessageBox.warning(self, "警告", "请先选择文件夹！")
            return
            
        self.log_text.clear()
        self.log_text.append("=== 预览重命名结果 ===")
        
        file_items = [self.file_list.item(i) for i in range(self.file_list.count())]
        if not file_items:
            self.log_text.append("没有文件可重命名")
            return
            
        start_index = self.start_index_spin.value()
        
        for i, item in enumerate(file_items):
            # 从显示文本中提取原始文件名
            original_text = item.text()
            old_name = original_text.split(" ")[0]  # 获取文件名部分
            
            new_name = self.generate_new_name(old_name, start_index + i)
            self.log_text.append(f"{i+1:03d}. {old_name} → {new_name}")
            
        self.rename_btn.setEnabled(True)
        self.log_text.append("预览完成，可以执行重命名")
        
    def execute_rename(self):
        """执行重命名操作"""
        if not self.folder_path:
            return
            
        reply = QMessageBox.question(self, "确认", "确定要执行重命名操作吗？此操作不可撤销！",
                                   QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
            
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(self.file_list.count())
        self.progress_bar.setValue(0)
        
        success_count = 0
        error_count = 0
        start_index = self.start_index_spin.value()
        
        self.log_text.append("=== 开始重命名 ===")
        
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            original_text = item.text()
            old_name = original_text.split(" ")[0]
            old_path = os.path.join(self.folder_path, old_name)
            
            new_name = self.generate_new_name(old_name, start_index + i)
            new_path = os.path.join(self.folder_path, new_name)
            
            try:
                os.rename(old_path, new_path)
                self.log_text.append(f"✓ 成功: {old_name} → {new_name}")
                success_count += 1
            except Exception as e:
                self.log_text.append(f"✗ 失败: {old_name} → {str(e)}")
                error_count += 1
                
            self.progress_bar.setValue(i + 1)
            QApplication.processEvents()  # 更新UI
            
        self.progress_bar.setVisible(False)
        self.log_text.append(f"=== 完成 ===")
        self.log_text.append(f"成功: {success_count}, 失败: {error_count}")
        
        # 更新文件列表
        self.update_file_list()
        self.rename_btn.setEnabled(False)

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # 使用Fusion样式
    
    # 设置字体
    font = QFont("Microsoft YaHei", 10)
    app.setFont(font)
    
    window = FileRenamerApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()