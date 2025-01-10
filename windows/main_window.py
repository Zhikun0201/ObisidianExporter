import os
import re
import subprocess
import sys

from PySide6.QtCore import Qt, QDir, QSettings, QFileInfo, QDirIterator, QCoreApplication
from PySide6.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QSplitter, QLabel, QWidget, QPushButton, \
    QSizePolicy, QTreeWidget, QTextBrowser, QFileDialog, QTreeView, QFileSystemModel, QApplication, \
    QListWidget, QListWidgetItem, QCheckBox, QMessageBox
import pypandoc

if hasattr(sys, '_MEIPASS'):
    # PyInstaller 在打包后，_MEIPASS 指向临时解压/运行目录
    base_path = sys._MEIPASS
    pandoc_path = os.path.join(base_path, 'pandoc/pandoc')
    pypandoc.__pandoc_path = pandoc_path

class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Obsidian 文档导出")
        self.setMinimumSize(400, 300)

        self.__selected_selected_files = []

        self.__init_layout()
        self.__init_button_bar()
        self.__init_file_tree()
        self.__init_preview()
        self.__init_config()

        self.__restore_settings()

        self.__connect_signals()

    def closeEvent(self, event):
        QSettings().setValue("window_size", self.size())
        # List items
        QSettings().setValue("only_include_list",
                             [self.only_include_list.item(i).text() for i in range(self.only_include_list.count())])
        QSettings().setValue("only_exclude_list",
                             [self._only_exclude_list.item(i).text() for i in range(self._only_exclude_list.count())])
        QSettings().setValue("hide_header_list",
                             [self.hide_header_list.item(i).text() for i in range(self.hide_header_list.count())])
        QSettings().setValue("hide_file_header", self.hide_file_header.isChecked())
        QSettings().setValue("hide_separator", self.hide_separator.isChecked())
        event.accept()

    def focusInEvent(self, event):
        self.__refresh_preview()
        event.accept()

    def __restore_settings(self):
        if QSettings().contains("window_size"):
            self.resize(QSettings().value("window_size"))
        if QSettings().contains("only_include_list"):
            for _item in QSettings().value("only_include_list"):
                _widget_item = QListWidgetItem(_item)
                _widget_item.setFlags(_widget_item.flags() | Qt.ItemFlag.ItemIsEditable)
                self.only_include_list.addItem(_widget_item)
        if QSettings().contains("only_exclude_list"):
            for _item in QSettings().value("only_exclude_list"):
                _widget_item = QListWidgetItem(_item)
                _widget_item.setFlags(_widget_item.flags() | Qt.ItemFlag.ItemIsEditable)
                self._only_exclude_list.addItem(_widget_item)
        if QSettings().contains("hide_header_list"):
            for _item in QSettings().value("hide_header_list"):
                _widget_item = QListWidgetItem(_item)
                _widget_item.setFlags(_widget_item.flags() | Qt.ItemFlag.ItemIsEditable)
                self.hide_header_list.addItem(_widget_item)
        if QSettings().contains("hide_file_header"):
            self.hide_file_header.setChecked(True if QSettings().value("hide_file_header") == "true" else False)
        if QSettings().contains("hide_separator"):
            self.hide_separator.setChecked(True if QSettings().value("hide_separator") == "true" else False)

    def __connect_signals(self):
        self._open_depo_button.clicked.connect(self.__open_depo)
        self._export_word_button.clicked.connect(self.__export_word)
        self._export_markdown_button.clicked.connect(self.__export_markdown)
        self._file_tree_view.clicked.connect(self.__update_selected_files)
        self._file_tree_model.rootPathChanged.connect(self.preview.setMarkdown(""))
        self.only_include_list.itemChanged.connect(self.__refresh_preview)
        self.only_include_list_add_button.clicked.connect(self.__add_include)
        self.only_include_list_remove_button.clicked.connect(self.__remove_include)
        self._only_exclude_list.itemChanged.connect(self.__refresh_preview)
        self.only_exclude_list_add_button.clicked.connect(self.__add_exclude)
        self.only_exclude_list_remove_button.clicked.connect(self.__remove_exclude)
        self.hide_header_list.itemChanged.connect(self.__refresh_preview)
        self.hide_header_list_add_button.clicked.connect(self.__add_hide_header)
        self.hide_header_list_remove_button.clicked.connect(self.__remove_hide_header)
        self.hide_file_header.stateChanged.connect(self.__refresh_preview)
        self.hide_separator.stateChanged.connect(self.__refresh_preview)

    def __init_layout(self):
        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)

        self.main_layout = QVBoxLayout()
        self.main_widget.setLayout(self.main_layout)
        # button bar
        self.button_bar_layout = QHBoxLayout()
        self.button_bar_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        self.main_layout.addLayout(self.button_bar_layout)
        # main func area
        self.main_func_layout = QHBoxLayout()
        self.main_layout.addLayout(self.main_func_layout)
        self.main_func_splitter = QSplitter()
        self.main_func_splitter.setChildrenCollapsible(False)
        self.main_func_layout.addWidget(self.main_func_splitter)
        # file tree
        self.file_tree_widget = QWidget()
        self.main_func_splitter.addWidget(self.file_tree_widget)
        self.file_tree_layout = QVBoxLayout()
        self.file_tree_widget.setLayout(self.file_tree_layout)
        # preview
        self.preview_widget = QWidget()
        self.main_func_splitter.addWidget(self.preview_widget)
        self.preview_layout = QVBoxLayout()
        self.preview_widget.setLayout(self.preview_layout)
        # config
        self.config_widget = QWidget()
        self.main_func_splitter.addWidget(self.config_widget)
        self.config_layout = QVBoxLayout()
        self.config_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.config_widget.setLayout(self.config_layout)

    def __init_button_bar(self):
        self._open_depo_button = QPushButton("打开 Obsidian Depo")
        self.button_bar_layout.addWidget(self._open_depo_button)

        _size_policy = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self._open_depo_button.setSizePolicy(_size_policy)

        self._export_word_button = QPushButton("导出 Word")
        self.button_bar_layout.addWidget(self._export_word_button)

        self._export_markdown_button = QPushButton("导出 Markdown")
        self.button_bar_layout.addWidget(self._export_markdown_button)

    def __init_file_tree(self):
        self._file_tree_model = QFileSystemModel()

        self._file_tree_view = QTreeView()
        self.file_tree_layout.addWidget(self._file_tree_view)
        self._file_tree_view.setModel(self._file_tree_model)

        self._file_tree_view.setHeaderHidden(True)
        self._file_tree_view.setColumnHidden(1, True)
        self._file_tree_view.setColumnHidden(2, True)
        self._file_tree_view.setColumnHidden(3, True)

        self._file_tree_view.setSelectionBehavior(QTreeWidget.SelectionBehavior.SelectRows)
        self._file_tree_view.setSelectionMode(QTreeView.SelectionMode.ExtendedSelection)

        self._file_tree_model.setReadOnly(True)

        if QSettings().contains("last_root_path"):
            self._file_tree_model.setRootPath(QSettings().value("last_root_path"))
            self._file_tree_view.setRootIndex(self._file_tree_model.index(QSettings().value("last_root_path")))

    def __init_preview(self):
        self.preview = QTextBrowser()
        self.preview_layout.addWidget(self.preview)

    def __init_config(self):
        self.only_include = QLabel("仅包含以下标题的内容：")
        self.config_layout.addWidget(self.only_include)
        self.only_include_list = QListWidget()
        self.config_layout.addWidget(self.only_include_list)
        self.only_include_list_button_group_layout = QHBoxLayout()
        self.config_layout.addLayout(self.only_include_list_button_group_layout)
        self.only_include_list_add_button = QPushButton("添加")
        self.only_include_list_button_group_layout.addWidget(self.only_include_list_add_button)
        self.only_include_list_remove_button = QPushButton("移除")
        self.only_include_list_button_group_layout.addWidget(self.only_include_list_remove_button)

        self.only_exclude = QLabel("不包含以下标题的内容：")
        self.config_layout.addWidget(self.only_exclude)
        self._only_exclude_list = QListWidget()
        self._only_exclude_list.setEditTriggers(QListWidget.EditTrigger.DoubleClicked)
        self.config_layout.addWidget(self._only_exclude_list)
        self.only_exclude_list_button_group_layout = QHBoxLayout()
        self.config_layout.addLayout(self.only_exclude_list_button_group_layout)
        self.only_exclude_list_add_button = QPushButton("添加")
        self.only_exclude_list_button_group_layout.addWidget(self.only_exclude_list_add_button)
        self.only_exclude_list_remove_button = QPushButton("移除")
        self.only_exclude_list_button_group_layout.addWidget(self.only_exclude_list_remove_button)

        self.hide_header = QLabel("隐藏以下标题：")
        self.config_layout.addWidget(self.hide_header)
        self.hide_header_list = QListWidget()
        self.config_layout.addWidget(self.hide_header_list)
        self.hide_header_list_button_group_layout = QHBoxLayout()
        self.config_layout.addLayout(self.hide_header_list_button_group_layout)
        self.hide_header_list_add_button = QPushButton("添加")
        self.hide_header_list_button_group_layout.addWidget(self.hide_header_list_add_button)
        self.hide_header_list_remove_button = QPushButton("移除")
        self.hide_header_list_button_group_layout.addWidget(self.hide_header_list_remove_button)

        self.hide_file_header = QCheckBox("隐藏文件标题")
        self.config_layout.addWidget(self.hide_file_header)
        self.hide_separator = QCheckBox("隐藏分隔线")
        self.config_layout.addWidget(self.hide_separator)

    def __open_depo(self):
        __dir_dialog = QFileDialog(self, "选择 Obsidian 仓库文件夹")
        __dir_dialog.setFileMode(QFileDialog.FileMode.Directory)
        __dir_dialog.setOption(QFileDialog.Option.ShowDirsOnly)
        if __dir_dialog.exec():
            __selected_dir = __dir_dialog.selectedFiles()
            if __selected_dir:
                root_path = __selected_dir[0]
                QSettings().setValue("last_root_path", root_path)
                if self._file_tree_model.setRootPath(root_path):
                    root_index = self._file_tree_model.index(root_path)
                    if root_index.isValid():
                        self._file_tree_view.setRootIndex(root_index)

    def __update_selected_files(self, index):
        cur_selected = self._file_tree_view.selectedIndexes()

        _pending_kill = []
        for _sel_index in self.__selected_selected_files:
            if _sel_index not in cur_selected:
                _pending_kill.append(_sel_index)
        for _kill in _pending_kill:
            self.__selected_selected_files.remove(_kill)

        if index not in self.__selected_selected_files:
            self.__selected_selected_files.append(index)

        self.__refresh_preview()

    def __add_include(self):
        _widget_item = QListWidgetItem("## 双击编辑标题，前缀 # 号数量表示标题级别")
        _widget_item.setFlags(_widget_item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.only_include_list.addItem(_widget_item)
        self.__refresh_preview()

    def __remove_include(self):
        for _item in self.only_include_list.selectedItems():
            self.only_include_list.takeItem(self.only_include_list.row(_item))
        self.__refresh_preview()

    def __add_exclude(self):
        _widget_item = QListWidgetItem("## 双击编辑标题，前缀 # 号数量表示标题级别")
        _widget_item.setFlags(_widget_item.flags() | Qt.ItemFlag.ItemIsEditable)
        self._only_exclude_list.addItem(_widget_item)
        self.__refresh_preview()

    def __remove_exclude(self):
        for _item in self._only_exclude_list.selectedItems():
            self._only_exclude_list.takeItem(self._only_exclude_list.row(_item))
        self.__refresh_preview()

    def __add_hide_header(self):
        _widget_item = QListWidgetItem("## 双击编辑标题，前缀 # 号数量表示标题级别")
        _widget_item.setFlags(_widget_item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.hide_header_list.addItem(_widget_item)
        self.__refresh_preview()

    def __remove_hide_header(self):
        for _item in self.hide_header_list.selectedItems():
            self.hide_header_list.takeItem(self.hide_header_list.row(_item))
        self.__refresh_preview()

    def __export_word(self):
        _saved_file, _ = QFileDialog(self, "选择导出 Word 的位置").getSaveFileName()
        if _saved_file:
            if not _saved_file.endswith(".docx"):
                _saved_file += ".docx"
            if convert_markdown_to_word(self.preview.toMarkdown(), _saved_file):
                if QMessageBox.information(self, "导出 Word", "导出 Word 成功！是否打开？",
                                           QMessageBox.StandardButton.Yes,
                                           QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                    if sys.platform == "win32":
                        os.startfile(_saved_file)
                    else:
                        opener = "open" if sys.platform == "darwin" else "xdg-open"
                        subprocess.call([opener, _saved_file])

    def __export_markdown(self):
        _saved_file, _ = QFileDialog(self, "选择导出 Markdown 的位置").getSaveFileName()
        if _saved_file:
            if not _saved_file.endswith(".md"):
                _saved_file += ".md"
            if save_markdown_file(self.preview.toMarkdown(), _saved_file):
                if QMessageBox.information(self, "导出 Markdown", "导出 Markdown 成功！是否打开？",
                                           QMessageBox.StandardButton.Yes,
                                           QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes:
                    if sys.platform == "win32":
                        os.startfile(_saved_file)
                    else:
                        opener = "open" if sys.platform == "darwin" else "xdg-open"
                        subprocess.call([opener, _saved_file])

    def __refresh_preview(self):
        files = []

        for _item in self.__selected_selected_files:
            _file_path = self._file_tree_model.filePath(_item)
            if QFileInfo(_file_path).isDir():
                f_dir = QDir(_file_path)
                it = QDirIterator(f_dir, QDirIterator.IteratorFlag.Subdirectories)
                while it.hasNext():
                    _file = it.next()
                    if QFileInfo(_file).isFile():
                        if _file not in files:
                            files.append(_file)
            else:
                if _file_path not in files:
                    files.append(_file_path)

        content = "---\n\n---\n\n"
        for _file in files:
            with open(_file, "r", encoding="utf-8") as f:
                f_read = f.read()

                # 去除文档开头的元数据
                f_read_split = f_read.split("---")
                if len(f_read_split) > 1 and f_read.startswith("---"):
                    f_read = f_read_split[2]

                # 仅包含标题
                include_list = []
                if self.only_include_list.count() > 0:
                    for i in range(self.only_include_list.count()):
                        target_header = self.only_include_list.item(i).text()
                        section = extract_section(f_read, target_header)
                        if section:
                            include_list.append(section)
                if include_list:
                    f_read = "\n\n".join(include_list)

                # 不包含标题
                exclude_list = []
                if self._only_exclude_list.count() > 0:
                    for i in range(self._only_exclude_list.count()):
                        target_header = self._only_exclude_list.item(i).text()
                        section = extract_section(f_read, target_header)
                        if section:
                            exclude_list.append(section)
                if exclude_list:
                    for section in exclude_list:
                        f_read = f_read.replace(section, "")

                # 替换 ![[inline]] 链接
                inline_link_pattern = r"!\[\[(.+?)\]\]"  # 匹配 ![[link]] 的正则
                inline_link_list = re.findall(inline_link_pattern, f_read)  # 获取所有匹配项
                inline_link_map = {}
                for inline_link in inline_link_list:
                    f_name, section_name = inline_link.split("#") if "#" in inline_link else (inline_link, "")
                    f_name = f_name + ".md" if not f_name.endswith(".md") else f_name
                    # 查找名称匹配的文件
                    _root_path = self._file_tree_model.rootPath()
                    j_it = QDirIterator(_root_path, QDirIterator.IteratorFlag.Subdirectories)
                    while j_it.hasNext():
                        _j_file = j_it.next()
                        if QFileInfo(_j_file).isFile() and QFileInfo(_j_file).fileName() == f_name:
                            with open(_j_file, "r", encoding="utf-8") as j:
                                if section_name:
                                    inline_link_section = extract_section(j.read(), section_name)
                                    if inline_link_section:
                                        inline_link_map[inline_link] = inline_link_section
                                else:
                                    inline_link_map[inline_link] = j.read()
                                break

                for inline_link, inline_content in inline_link_map.items():
                    f_read = f_read.replace(f"![[{inline_link}]]", inline_content)

                # 替换 [[link]] 链接
                pattern = r"\[\[([^\[\]|]+)\|?([^\[\]]*)\]\]"
                f_read = re.sub(pattern, lambda m: m.group(2) if m.group(2) else m.group(1), f_read)

                # 隐藏标题
                if self.hide_header_list.count() > 0:
                    for i in range(self.hide_header_list.count()):
                        target_header = self.hide_header_list.item(i).text()
                        f_read = replace_section_title(f_read, target_header, "")

                content += "# " + QFileInfo(_file).fileName().rstrip(
                    ".md") + "\n\n" if not self.hide_file_header.isChecked() else ""
                content += f_read + "\n\n"
                content += "---\n\n" if not self.hide_separator.isChecked() else ""

        # content += "---" if not self.hide_separator.isChecked() else ""
        self.preview.setMarkdown(content)


def replace_section_title(markdown_text, old_title, new_title):
    # 匹配标题的正则，忽略标题的级别（任意数量的 #）
    header_level = old_title.count("#")  # 获取目标标题的级别
    if header_level != 0:
        pattern = rf"(^{re.escape(old_title)}\s*\n)"
        # rf"(^#+\s*{re.escape(old_title)})(\s*\n)"
    else:
        pattern = rf"(^#+\s*{re.escape(old_title)}\s*\n)"
    # 替换标题，不影响内容
    replacement = rf"{new_title}"  # 替换为新标题（可调整级别）
    # 使用 re.MULTILINE 保证 ^ 匹配行首
    updated_text = re.sub(pattern, replacement, markdown_text, flags=re.MULTILINE)
    return updated_text


def extract_section(markdown_text, target_header):
    # 构建目标标题的正则模式，包括内容及子标题
    # 捕获从目标标题开始到下一个同级或更高级标题为止的所有内容
    header_level = target_header.count("#")  # 获取目标标题的级别
    if header_level != 0:
        pattern = rf"(^{re.escape(target_header)}\s*\n)(.*?)(?=^#{{1,{header_level}}} |\Z)"
    else:
        pattern = rf"(^#+\s*{re.escape(target_header)}\s*\n)(.*?)(?=^#|\Z)"
    # 使用 re.DOTALL 使 `.*?` 跨行匹配，并使用 re.MULTILINE 处理多行
    match = re.search(pattern, markdown_text, re.MULTILINE | re.DOTALL)

    if match:
        return match.group(0)  # 返回匹配到的内容
    else:
        return None


# Markdown 转 Word
def convert_markdown_to_word(markdown_content, output_path):
    if pypandoc.get_pandoc_path() is None:
        pypandoc.download_pandoc()
    markdown_content = markdown_content.replace("- - -", "")
    pypandoc.convert_text(markdown_content, "docx", format="md", outputfile=output_path)
    return True


# 保存 Markdown 文件
def save_markdown_file(markdown_content, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(markdown_content)
    return True


if __name__ == "__main__":
    app = QApplication(sys.argv)

    QCoreApplication.setOrganizationName("Obsidian Exporter")
    QCoreApplication.setApplicationName("Obsidian Exporter")

    main_window = MainWindow()
    main_window.show()

    app.exec()
