import sys
import os
from arrow import now
import requests
import xml.etree.ElementTree as ET
from pathlib import Path

import matplotlib as mpl
import matplotlib.pyplot as plt
mpl.rcParams.update({
    "font.size": 8
})

from datetime import datetime, timedelta
from scipy.interpolate import interp1d
import numpy as np
import time
from collections import Counter
from scipy.signal import savgol_filter

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx2pdf import convert

import lib_vspec as vsp
import pyVOBeSS as bess


from PySide6.QtUiTools import QUiLoader
from PySide6.QtWidgets import QApplication,QMainWindow,QFileDialog,QMessageBox,QWidget, QDialog
from PySide6.QtCore import QFile, QLocale, QIODevice, Qt, QDate, QSettings, QTranslator, Signal, QObject
from PySide6 import QtGui

from mplcanvas import MplCanvas

original_stdout = sys.stdout

class main_wnd_UI(QMainWindow) :
    def __init__(self):
        super(main_wnd_UI, self).__init__()

        self.version ="0.1"
        
        #fichier GUI par Qt Designer
        loader = QUiLoader()
        loader.registerCustomWidget(MplCanvas)
        ui_file_name=resource_path('BeSS_report.ui')
        ui_file = QFile(ui_file_name)
        
        if not ui_file.open(QIODevice.ReadOnly):
            print(f"Cannot open {ui_file_name}: {ui_file.errorString()}")
            sys.exit(-1)
        
        self.ui = loader.load(ui_file)
        ui_file.close()

        # redirection
        self.stdout_redirector = StdoutRedirector()
        self.stdout_redirector.new_text.connect(self.add_text)
        sys.stdout = self.stdout_redirector
        print("BeSS report version : "+ self.version)
        print(' ')

        # initialisation des widgets
        self.ui.month_txt.setText(str(datetime.now().month - 1))
        self.ui.year_txt.setText(str(datetime.now().year))

        self.ui.month_spc_txt.setText(str(datetime.now().month - 1))
        #self.ui.date_fin_ctrl.setDate(QDate.currentDate())

        self.ui.month_evo_txt.setText(str(datetime.now().month - 1))
        self.ui.nb_max_txt.setText('10')


        # connect signals to slots
        self.ui.ok_report_btn.clicked.connect(self.ok_report_clicked)
        self.ui.ok_spc_btn.clicked.connect(self.ok_obj_allspec_clicked)
        self.ui.ok_evo_btn.clicked.connect(self.ok_obj_evo_clicked)

        self.ui.export_btn.clicked.connect(self.export_clicked)
        self.ui.cls_btn.clicked.connect(self.clear_clicked)
        self.ui.exit_btn.clicked.connect(self.exit_clicked)

    def show(self) :
        self.ui.show()
        
    def closeEvent (self,event):
        sys.stdout=original_stdout
        #QApplication.restoreOverrideCursor()

    def exit_clicked(self) :
        sys.stdout=original_stdout
        plt.close('all')
        print('exit')
        QApplication.instance().quit()
        app.quit()

    def add_text(self, text) :
        self.ui.console_txt.append(text)
        self.ui.console_txt.moveCursor(QtGui.QTextCursor.End)

    def export_clicked (self):
        log_file = Path("log.txt")
        with open(log_file, "w", encoding="utf-8") as f:
            f.write(self.ui.console_txt.toPlainText())

    def clear_clicked(self) :
        self.ui.console_txt.clear()
        print("BeSS report version : "+ self.version)
        print(' ')

    def ok_report_clicked(self) :
        mois = int(self.ui.month_txt.text())
        year = int(self.ui.year_txt.text())
        flag_metrics = self.ui.metric_chk.isChecked()
        flag_with_novar = self.ui.novar_chk.isChecked()
        flag_noreport = self.ui.noreport_chk.isChecked()
        nb_max = 10

        flag = 2
        bess.main_dispatch(flag, mois, year, object_name=None,date_deb=None, date_fin=None, nb_max=nb_max,
                           flag_metrics=flag_metrics, flag_with_novar=flag_with_novar, flag_noreport=flag_noreport)
        
    def ok_obj_allspec_clicked(self) :
        object_name=self.ui.obj_spc_txt.text()
        mois = int(self.ui.month_spc_txt.text())
        nb_max = int(self.ui.nb_max_txt.text())
        flag_composer =  self.ui.composer_chk.isChecked()
        #date_deb = self.ui.date_deb_ctrl.date().toPython()
        #date_fin = self.ui.date_fin_ctrl.date().toPython()

        # Download tous les spectres entre date_deb et date_fin pour un objet
        save_dir = Path(__file__).resolve().parent / "BeSS_VO"

        # récupère le mois et l'année courante
        year_now = datetime.now().year
        
        flag = 1 # pour all spectra
        bess.main_dispatch(flag, mois, year_now, object_name,date_deb=None, date_fin=None, nb_max=nb_max, flag_composer=flag_composer)
        

    def ok_obj_evo_clicked(self) :
        object_name=self.ui.obj_evo_txt.text()
        mois = int(self.ui.month_evo_txt.text())
        flag_metrics = self.ui.evo_metric_chk.isChecked()

        # récupère l'année courante
        year_now = datetime.now().year

        flag = 0 # pour evolution
        bess.main_dispatch(flag, mois, year_now, object_name,date_deb=None, date_fin=None, flag_metrics=flag_metrics)



    

# ----------------------------------------------------------------------------
# class redirection console to textEdit
#-----------------------------------------------------------------------------
class StdoutRedirector(QObject):
    new_text = Signal(str)

    def write(self, text):
        if text==' ' :
            self.new_text.emit(text)
        else :
            if text.strip() :
                self.new_text.emit(text)
        
        
    def flush(self) :
        pass


# ------------------------------------------------------------------------------
# fonctions
# -----------------------------------------------------------------------------

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

def data_path(relative_path):
    """ Get path to exe, works for dev and for PyInstaller """
    data_path = os.path.abspath(os.path.join(os.path.dirname(sys.argv[0]), relative_path))
    return data_path


#*****************************************************************************
# main app Qt
#*****************************************************************************  

if __name__ == "__main__":
    
    loader = QUiLoader()    
       
    # pour eviter de devoir tuer app qt sous spyder
    app = QApplication.instance() 
    if not app:
        app = QApplication(sys.argv)
    else:
        app = QApplication.instance()
    
    app.setStyle('fusion') # pour forcer un beau look sur Mac
 
    my_wnd_class=main_wnd_UI() 
    my_wnd_class.show()
    
    sys.exit(app.exec())

