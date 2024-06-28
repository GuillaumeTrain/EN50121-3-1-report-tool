from PySide6.QtGui import QColor
from PySide6.QtWidgets import QMainWindow, QApplication, QVBoxLayout, QWidget, QPushButton, QLineEdit, QHBoxLayout, \
    QFileDialog, QTreeView, QFileSystemModel, QSplitter, QListView, QListWidgetItem, QLabel, QCheckBox, QColorDialog, \
    QListWidget, QScrollArea, QStyle, QSizePolicy, QComboBox
from PySide6.QtCore import QFile, QTextStream, QDir, QModelIndex
from PySide6.QtCore import Qt
from PySide6.QtCore import Signal

import os


class FileBrowser(QWidget):
    def __init__(self):
        super().__init__()
        self.current_path = ""
        self.resize(300, 800)
        self.fileBrowserWidget = QWidget(self)

        self.dirmodel = QFileSystemModel()
        # Don't show files, just folders
        self.dirmodel.setFilter(QDir.NoDot | QDir.AllDirs)
        self.folder_view = QTreeView(parent=self)
        self.folder_view.setModel(self.dirmodel)
        self.folder_view.clicked[QModelIndex].connect(self.clicked)

        # Don't show columns for size, file type, and last modified
        self.folder_view.setHeaderHidden(True)
        self.folder_view.hideColumn(1)
        self.folder_view.hideColumn(2)
        self.folder_view.hideColumn(3)

        self.selectionModel = self.folder_view.selectionModel()
        self.filemodel = QFileSystemModel()
        # Don't show folders, just files
        self.filemodel.setFilter(QDir.NoDotAndDotDot | QDir.Files)
        self.file_view = QListView(parent=self)
        self.file_view.setModel(self.filemodel)
        self.set_path()
        # self.file_view.clicked.connect(self.fileselected)

        self.splitter_filebrowser = QSplitter()
        self.splitter_filebrowser.addWidget(self.folder_view)
        self.splitter_filebrowser.addWidget(self.file_view)
        # self.splitter_filebrowser.setStretchFactor(0, 2)
        # self.splitter_filebrowser.setStretchFactor(1, 4)

        self.refreshbutton = QPushButton("Refresh")
        self.layout = QVBoxLayout()
        self.fileBrowserWidget.setLayout(self.layout)
        self.hbox1 = QHBoxLayout()
        self.layout.addLayout(self.hbox1)
        self.hbox1.addWidget(self.splitter_filebrowser)
        self.hbox2 = QHBoxLayout()
        self.hbox2.addWidget(self.refreshbutton)
        self.layout.addLayout(self.hbox2)

        self.setLayout(self.layout)

        self.refreshbutton.clicked.connect(self.refresh)

    def set_path(self, path=""):  # add a path parameter
        self.current_path = path  # update the current path
        self.dirmodel.setRootPath(path)
        self.filemodel.setRootPath(path)

    def refresh(self):
        self.set_path(self.current_path)  # refresh the current path
        self.folder_view.close()
        self.dirmodel.setFilter(QDir.NoDot | QDir.AllDirs)
        self.folder_view = QTreeView(parent=self);
        self.folder_view.setModel(self.dirmodel)
        self.folder_view.clicked[QModelIndex].connect(self.clicked)
        self.folder_view.setHeaderHidden(True)
        self.folder_view.hideColumn(1)
        self.folder_view.hideColumn(2)
        self.folder_view.hideColumn(3)
        self.selectionModel = self.folder_view.selectionModel()
        self.splitter_filebrowser.insertWidget(0, self.folder_view)

    def clicked(self, index):
        # get selected path of folder_view
        index = self.selectionModel.currentIndex()
        dir_path = self.dirmodel.filePath(index)
        self.set_path(dir_path)  # set the path when a directory is clicked
        self.file_view.setRootIndex(self.filemodel.index(dir_path))


class QListPlotItems(QScrollArea):
    def __init__(self, root):
        super().__init__()
        # Create a QVBoxLayout inside the widget
        self.root = root
        self.layout = QVBoxLayout(self)
        self.layout.setAlignment(Qt.AlignTop)
        self.items = []
        self.selected_item = None
        self.selected_index = None
        self.layout.setSpacing(0)
        self.layout.setContentsMargins(0, 0, 0, 0)
        # self.layout.setSpacing(0)
        # self.layout.addStretch()

    def add_item(self, filepath):
        item = QMeasItem(self)
        item.set_file(filepath)
        self.items.append(item)
        self.layout.addWidget(item)
        # item.selected.connect(self.select_item)

    def select_item(self, item):
        if self.selected_item:
            self.selected_item.deselect_item()
        self.selected_item = item
        self.selected_item.select_item()
        self.root.current_list = self
        #print(self.root.current_list)
        try:
            index = self.items.index(item)
            self.selected_index = index
            #print(index)
        except Exception:
            pass
            # print("echec de la selection")

    def delete_selected_item(self):
        if self.selected_item:
            self.layout.removeWidget(self.selected_item)

            item = self.items[self.selected_index]
            # print(item)
            self.selected_item.deleteLater()
            self.items.remove(item)
            self.selected_item = None
            self.selected_index = None

    def get_selected_index(self):

        return self.selected_index

    def get_selected_item(self):

        return self.selected_item

    def get_item(self, index: int):

        return self.items[index]

    def get_item_filepath(self, index: int):
        item = self.get_item(index)
        fp = item.get_filepath(item)

        return fp

    def countitems(self):
        return len(self.items)

    def get_filepath_list(self):

        fp_list = []

        for item in self.items:

            fp_list.append(item.get_filepath())

        return fp_list

    def get_item_color(self, index):
        return self.items[index].get_color()


class QMeasItem(QWidget):

    def __init__(self, root: QListPlotItems):
        super().__init__()
        self.root = root

        #print(root.items)
        self.selected = Signal(object)
        self.label = QLabel()
        self.layout = QHBoxLayout()
        self.layout.setSpacing(0)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.label.setContentsMargins(0, 0, 0, 0)

        self.filepath = None
        self.trace_nbr = QComboBox()
        self.trace_nbr.addItems(["1", "2", "3", "4", "5", "6"])
        self.color = QPushButton(" ")
        self.color.setContentsMargins(0, 0, 0, 0)
        self.color.setStyleSheet("border-radius: 7px;")
        self.button_color: QColor = QColor(92, 85, 233)
        self.checkbox = QCheckBox()
        self.data = None

        self.layout.addWidget(self.label)
        self.layout.addWidget(self.color)
        self.layout.addWidget(self.checkbox)
        self.setLayout(self.layout)
        self.color.clicked.connect(self.open_color_dialog)
        self.layout.addWidget(self.label, stretch=3)  # Give more space to the label
        self.layout.addWidget(self.trace_nbr, stretch=1)  # Give less space to the color button
        self.layout.addWidget(self.color, stretch=1)  # Give less space to the color button
        self.layout.addWidget(self.checkbox, stretch=1)
        self.label.mousePressEvent = lambda event: self.root.select_item(self)
        # self.clicked.connect(self.select_item)

    def is_checked(self):
        return self.checkbox.isChecked()

    def open_color_dialog(self):
        self.color_dialog = QColorDialog(self)
        self.color_dialog.colorSelected.connect(self.change_button_color)
        self.color_dialog.show()

    def change_button_color(self, color: QColor):
        self.color.setStyleSheet(f"border-radius: 7px;background-color: {color.name()};")
        self.button_color = color

    def get_color(self):
        return self.button_color

    def set_filename(self, filename):
        self.label.setText(filename)

    def get_filename(self):
        return self.label.text()
    def set_filepath(self, filepath):
        self.filepath = filepath

    def get_filepath(self):
        return self.filepath

    def get_filename_from_path(self, filepath):
        return os.path.basename(filepath)

    def set_file(self, filepath):
        self.set_filepath(filepath)
        #filename = os.path.basename(filepath)
        # print(filename)
        # self.set_filename(filename)
        filename_without_ext = os.path.splitext(os.path.basename(filepath))[0]
        # print(filename_without_ext)
        self.set_filename(filename_without_ext)

    def select_item(self):
        self.label.setStyleSheet("background-color: yellow;")
        # self.setStyleSheet("background-color: yellow;")

        # self.selected.emit(self)

    def deselect_item(self):
        self.label.setStyleSheet("background-color: white;")
        # self.setStyleSheet("background-color: white;")
