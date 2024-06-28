from PySide6.QtWidgets import QMainWindow, QApplication, QVBoxLayout, QWidget, QPushButton, QLineEdit, QHBoxLayout,QFileDialog, QListWidget, QLabel, QCheckBox, QComboBox, QTabWidget
import pyqtgraph as pg
import sys
from csvreader import FileReader
from filebrowzer import FileBrowser as FB, QMeasItem
from filebrowzer import QListPlotItems as PL
from functools import reduce, partial
import dataprocessor as DP


class MainWindow(QMainWindow):

    def __init__(self):
        # Initialisation des attributs
        self.data_processor: DP.DataProcessor = None
        super(MainWindow, self).__init__()
        self.alldataframes = None
        self.xmin = 0
        self.xmax = 0

        # Création des widgets principaux
        main_widget = QWidget()
        measurement_widget = QWidget()
        self.analysis_widget = QWidget()
        file_manager_widget = QWidget()
        file_manager_widget.setMaximumWidth(500)

        # Configuration du layout principal
        main_layout = QVBoxLayout()
        main_layout.addWidget(measurement_widget)
        main_layout.addWidget(self.analysis_widget)
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)

        # Configuration du layout des mesures
        measurement_layout = QHBoxLayout()
        measurement_layout.addWidget(file_manager_widget)
        measurement_widget.setLayout(measurement_layout)

        # Configuration du layout d'analyse
        self.analysis_layout = QHBoxLayout()
        self.analysis_widget.setLayout(self.analysis_layout)

        self.extra_analysis = QTabWidget()
        self.ambiant_analysis_widget = QWidget()
        self.ambiant_analysis_layout = QHBoxLayout()
        self.ambiant_analysis_widget.setLayout(self.ambiant_analysis_layout)
        self.data_analysis_widget = QWidget()
        self.data_analysis_layout = QVBoxLayout()
        self.data_analysis_widget.setLayout(self.data_analysis_layout)
        self.extra_analysis.addTab(self.ambiant_analysis_widget, 'Ambient analysis')
        self.extra_analysis.addTab(self.data_analysis_widget,"Data Table")
        self.analysis_layout.addWidget(self.extra_analysis)
        self.data_table_processor = DP.DataTableProcessor(self)


        # Création et configuration des widgets d'analyse d'ambiant
        self.ambient_postprocessor = DP.AmbientPostProcessor(self)
        self.ambient_postprocessor.ambient_analysis_button.setEnabled(False)
        self.meas_postprocessor = DP.MeasPostProcessor(self)
#        self.meas_postprocessor.meas_analysis_button.setEnabled(False)

        # Configuration du layout du gestionnaire de fichiers
        file_manager_layout = QVBoxLayout()
        file_manager_widget.setLayout(file_manager_layout)

        # Création et configuration du widget de graphique
        self.graphWidget = pg.PlotWidget()
        self.graphWidget.sigXRangeChanged.connect(self.setYRange)
        self.graphWidget.setBackground('w')

        # Configuration du layout du graphique
        graph_layout = QVBoxLayout()
        graph_layout.addWidget(self.graphWidget)
        measurement_layout.addLayout(graph_layout)

        # Création et configuration du widget de dialogue de fichier
        self.file_diag = FB()
        self.file_diag.file_view.clicked.connect(self.fileselected)
        self.path_lineEdit = QLineEdit()
        file_manager_layout.addWidget(self.file_diag)
        file_manager_layout.addWidget(self.path_lineEdit)

        # Création et configuration des boutons et des listes
        self.ambiant_list = PL(self)
        self.current_list = self.ambiant_list
        self.ambient_button = QPushButton("Ambient")
        self.ambient_button.clicked.connect(self.on_click_add_to_ambient)
        to_right_button = QPushButton(">>")
        to_right_button.clicked.connect(self.to_right_button_clicked)

        self.meas_list = PL(self)
        self.meas_button = QPushButton("Measurements")
        self.meas_button.clicked.connect(self.on_click_add_to_meas)
        to_left_button = QPushButton("<<")
        to_left_button.clicked.connect(self.to_left_button_clicked)

        self.limit_list = PL(self)
        self.limit_button = QPushButton("Limit")
        self.limit_button.clicked.connect(self.on_click_add_to_limit)
        del_limit_button = QPushButton("delete")
        del_limit_button.clicked.connect(self.del_button_clicked)

        refresh_plt_button = QPushButton("Refresh Plot")
        refresh_plt_button.clicked.connect(self.refresh_plot)
        self.Yaxistype_qcombobox = QComboBox()
        self.Yaxistype_qcombobox.addItems(["dBµV/m","dBµA/m","dBµT","mT","V/m","A/m"])
        self.Yaxistype_qcombobox.currentIndexChanged.connect(self.update_axis_title)
        self.transducer_Label = QLabel("Transducer factor (dB) : ")
        self.transducer_factor = QLineEdit("0")
        self.transducer_help_Label = QLabel("Transducer factor (dB) : (if the transducer is an attenuator,\n the factor must be negative), the transducer factor is the difference between the transducer and the antenna factor will be applied\n for each measurements files")


        # Configuration des layouts des listes
        vambient_box = QVBoxLayout()
        vambient_box.addWidget(self.ambient_button)
        vambient_box.addWidget(to_right_button)
        vambient_box.addWidget(self.ambiant_list)

        vmeas_box = QVBoxLayout()
        vmeas_box.addWidget(self.meas_button)
        vmeas_box.addWidget(to_left_button)
        vmeas_box.addWidget(self.meas_list)

        vlimit_box = QVBoxLayout()
        vlimit_box.addWidget(self.limit_button)
        vlimit_box.addWidget(del_limit_button)
        vlimit_box.addWidget(self.limit_list)

        # Configuration du layout du gestionnaire de mesures
        meas_manager_layout = QHBoxLayout()
        meas_manager_layout.addLayout(vambient_box)
        meas_manager_layout.addLayout(vmeas_box)
        meas_manager_layout.addLayout(vlimit_box)
        file_manager_layout.addLayout(meas_manager_layout)
        file_manager_layout.addWidget(QLabel("Y Unit : "))
        file_manager_layout.addWidget(self.Yaxistype_qcombobox)
        file_manager_layout.addWidget(self.transducer_Label)
        file_manager_layout.addWidget(self.transducer_factor)
        file_manager_layout.addWidget(self.transducer_help_Label)
        file_manager_layout.addWidget(refresh_plt_button)

        self.graphWidget.getPlotItem().getAxis('left').setLabel('Level', units=self.Yaxistype_qcombobox.currentText())
        self.graphWidget.getPlotItem().getAxis('bottom').setLabel('Frequency', units='Hz')

    def update_axis_title(self):
        self.graphWidget.getPlotItem().getAxis('left').setLabel('Level', units=self.Yaxistype_qcombobox.currentText())

    def setYRange(self):

        self.graphWidget.enableAutoRange(axis='y')
        self.graphWidget.getPlotItem().getAxis('bottom').setTicks(None)

    def on_list_clicked(self, selectedlist):
        self.current_list = selectedlist
        print("selectedlist", selectedlist)
        print("current_list", self.current_list)

    def refresh_plot(self):
        self.graphWidget.clear()
        self.graphWidget.setLogMode(x=True, y=False)
        self.data_processor: DP.DataProcessor = DP.DataProcessor(self.ambiant_list, self.meas_list, self.limit_list,self)
        self.data_processor.set_transducer_factor(float(self.transducer_factor.text()))
        self.data_processor.apply_transducer_factor()
        self.graphWidget.addLegend()
        curve = self.data_processor.get_merged_ambient()
        if curve is not None:
            # print(curve.data.values)
            # print("test data ambient:{}".format(curve.data['Frequency']))
            # print("test data ambient:{}".format(curve.data['Field']))
            # print("test data ambient:{}".format(curve.color.getRgb()))
            self.graphWidget.plot(curve.data['Frequency'], curve.data['Field'], pen=curve.color.getRgb(), name=curve.name)
        meas_list = self.data_processor.get_meas_list()
        if meas_list is not None:
            for curve in meas_list:
                # print(curve.data.values)
                # print("test data ambient:{}".format(curve.data['Frequency']))
                # print("test data ambient:{}".format(curve.data['Field']))
                # print("test data ambient:{}".format(curve.color.getRgb()))
                self.graphWidget.plot(curve.data['Frequency'], curve.data['Field'], pen=curve.color.getRgb(), name = curve.name)
        lim_list = self.data_processor.get_limit_list()
        if lim_list is not None:
            for curve in lim_list:
                # print(curve.data.values)
                # print("test data ambient:{}".format(curve.data['Frequency']))
                # print("test data ambient:{}".format(curve.data['Field']))
                # print("test data ambient:{}".format(curve.color.getRgb()))
                self.graphWidget.plot(curve.data['Frequency'], curve.data['Field'], pen=curve.color.getRgb(), name = curve.name)
        # print("ok")
        self.graphWidget.enableAutoRange(axis='y')
        self.graphWidget.getPlotItem().showGrid(x=True, y=True, alpha=0.7)
        self.ambient_postprocessor.ambient_analysis_button.setEnabled(True)
        self.meas_postprocessor.meas_post_analysis_plot_button.setEnabled(True)
        self.graphWidget.getPlotItem().getAxis('left').setLabel('Level', units=self.Yaxistype_qcombobox.currentText())

    def del_button_clicked(self):
        current_index = self.current_list.selected_index
        # print(current_index)
        filepath = self.current_list.items[current_index].get_filepath()

        # print(filepath)
        self.current_list.delete_selected_item()

    def to_right_button_clicked(self):
        target_list = self.meas_list

        if self.current_list is self.meas_list:
            target_list = self.limit_list
        else:
            pass
        current_item = self.current_list.get_selected_item()
        filepath = current_item.get_filepath()
        try:
            print(filepath)
            target_list.add_item(filepath)
            self.current_list.delete_selected_item()
        except Exception:
            print("aucun item selectionne dans l'une des liste")

    def to_left_button_clicked(self):
        target_list = self.meas_list
        if self.current_list is self.meas_list:
            target_list = self.ambiant_list
        else:
            pass
        current_item = self.current_list.get_selected_item()
        filepath = current_item.get_filepath()
        try:
            print(filepath)
            target_list.add_item(filepath)
            self.current_list.delete_selected_item()
        except Exception:
            print("aucun item selectionne dans l'une des liste")

    def to_amb_button_clicked(self):
        meas_index = self.meas_list.currentIndex()
        #print(meas_index.row())
        filepath = self.meas_list.item(meas_index.row())
        #print(filepath.text())
        self.ambiant_list.addItem(filepath.text())
        self.meas_list.takeItem(meas_index.row())

    def on_click_add_to_ambient(self):
        self.ambiant_list.add_item(self.path_lineEdit.text())

    def on_click_add_to_meas(self):
        self.meas_list.add_item(self.path_lineEdit.text())

    def on_click_add_to_limit(self):
        self.limit_list.add_item(self.path_lineEdit.text())

    def fileselected(self, index):
        file_path = self.file_diag.filemodel.filePath(index)
        self.path_lineEdit.setText(file_path)
        #print(file_path)
        # file_path = self.filemodel.filePath(index)
        # print(file_path)
        return file_path


app = QApplication(sys.argv)
w = MainWindow()
w.show()
FileReader()
with open("Toolery.qss", "r") as f:
    _style = f.read()
    app.setStyleSheet(_style)
app.exec()
