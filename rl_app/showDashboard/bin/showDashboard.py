# -*- coding: utf-8 -*-

import os
import sys
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from PySide2.QtGui import QMouseEvent ,Qt ,QCursor , QPalette
from PySide2 import QtCore, QtWidgets ,QtGui

import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
import numpy as np

import cgt_info as cgt
import pipeline_plot_chart
import artist_plot_chart
import project_pie_chart

class dashboard(QWidget):
    
    # proj_signal = QtCore.Signal()
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Project DashBoard')
        
        self.setup_ui()
        
    def setup_ui(self):
        self.resize(1200,800)
        
        # palette = self.palette()
        # palette.setColor(QPalette.Background,QColor(75,77,83))
        # self.setPalette(palette)
        # self.setAutoFillBackground(True)
        
        self.combox_project = QComboBox()
        self.get_active_project()
        self.combox_project.setCurrentIndex(-1)
        
        self.radbtn_task_status = QRadioButton('Task Status')
        self.radbtn_task_status.setChecked(1)
        self.radbtn_commissioner_status = QRadioButton('Commissioner Status')
        self.radbtn_commissioner_status.setChecked(0)
        
        self.btn_showDashboard = QPushButton('Show Dashboard')
        
        self.tree_sequence = QTreeWidget()
        self.tree_sequence.setHeaderHidden(1)
        # self.tree_sequence.setStyleSheet("background-color:#4b4d53")
        
        self.lab_pipeline_total = QLabel('Pipeline Total')
        self.draw_pipeline = pipeline_plot_chart.draw_chart()
        
        self.lab_artist_total = QLabel('Artist Total')
        self.draw_artist = artist_plot_chart.draw_chart()
        
        self.lab_project_total = QLabel('Project Total')
        self.draw_project = project_pie_chart.draw_chart()
        
        self.status_lay = QHBoxLayout()
        self.status_lay.addWidget(self.radbtn_task_status)
        self.status_lay.addWidget(self.radbtn_commissioner_status)
        
        self.project_lay = QVBoxLayout()
        self.project_lay.addWidget(self.combox_project)
        self.project_lay.addLayout(self.status_lay)
        self.project_lay.addWidget(self.tree_sequence)
        
        self.pipeline_plot_chart_lay = QVBoxLayout()
        # self.pipeline_plot_chart_lay.addWidget(self.lab_pipeline_total)
        self.pipeline_plot_chart_lay.addWidget(self.draw_pipeline)
        
        self.artist_plot_chart_lay = QVBoxLayout()
        # self.artist_plot_chart_lay.addWidget(self.lab_artist_total)
        self.artist_plot_chart_lay.addWidget(self.draw_artist)
        
        self.project_pie_chart_lay = QVBoxLayout()
        # self.project_pie_chart_lay.addWidget(self.lab_project_total)
        self.project_pie_chart_lay.addWidget(self.draw_project)
        
        # self.dashboard_lay = QGridLayout()
        # self.dashboard_lay.addLayout(self.project_lay,0,0,3,1)
        # self.dashboard_lay.addWidget(self.btn_showDashboard,0,4,1,1)
        # self.dashboard_lay.addLayout(self.pipeline_plot_chart_lay,1,1,1,2)
        # self.dashboard_lay.addLayout(self.project_pie_chart_lay,1,3,1,2)
        # self.dashboard_lay.addLayout(self.artist_plot_chart_lay,2,1,1,4)
        
        self.main_lay = QGridLayout()
        self.main_lay.addWidget(self.btn_showDashboard,0,4,1,1)
        self.main_lay.addLayout(self.pipeline_plot_chart_lay,1,1,1,2)
        self.main_lay.addLayout(self.project_pie_chart_lay,1,3,1,2)
        self.main_lay.addLayout(self.artist_plot_chart_lay,2,1,1,4)
        
        self.dashboard_lay = QFormLayout()
        self.dashboard_lay.setLayout(0,QFormLayout.LabelRole,self.project_lay)
        self.dashboard_lay.setLayout(0,QFormLayout.FieldRole,self.main_lay)
        
        self.setLayout(self.dashboard_lay)
        
        self.combox_project.currentTextChanged.connect(self.get_current_project)
        self.btn_showDashboard.clicked.connect(self.show_dashboard)
        
    def get_active_project(self):
        project_data = cgt.project.get_active_project_dict()
        project_name = [proj_name for proj_name in project_data]
            
        project_name.sort()
        self.combox_project.addItems(project_name)
        
    def get_select_seq(self):
        seq_list = []
        seq_it = QTreeWidgetItemIterator(self.tree_sequence)
        while seq_it.value():
            if seq_it.value().isDisabled() == False and seq_it.value().checkState(0) == Qt.Checked:
                seq_list.append(seq_it.value().text(0))
            seq_it += 1
            
        return seq_list
        
    def get_current_project(self):
        current_project = self.combox_project.currentText()
        seq = cgt.project.get_seq(current_project)
        seq.sort()
        
        self.tree_sequence.clear()
        for seq_index in seq:
            seq_item = QTreeWidgetItem(self.tree_sequence)
            seq_item.setText(0,seq_index)
            seq_item.setExpanded(1)
            seq_item.setFlags(Qt.ItemFlag.ItemIsAutoTristate | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            if seq_index == 'dev' or seq_index == 'prod':
                seq_item.setCheckState(0,Qt.Unchecked)
            else:
                seq_item.setCheckState(0,Qt.Checked)
                
    def draw_pipeline_chart(self):
        current_project = self.combox_project.currentText()
        seq = self.get_select_seq()

        del_pipeline_lay = self.pipeline_plot_chart_lay.takeAt(0)
        del_pipeline_lay.widget().deleteLater()
        if self.radbtn_task_status.isChecked():
            pipeline_darw_plot = pipeline_plot_chart.draw_chart(current_project,seq,'task.status')
        elif self.radbtn_commissioner_status.isChecked():
            pipeline_darw_plot = pipeline_plot_chart.draw_chart(current_project,seq,'task.sup_review')
        self.pipeline_plot_chart_lay.addWidget(pipeline_darw_plot)
        
    def draw_artist_chart(self):
        current_project = self.combox_project.currentText()
        seq = self.get_select_seq()
        
        del_artist_lay = self.artist_plot_chart_lay.takeAt(0)
        del_artist_lay.widget().deleteLater()
        if self.radbtn_task_status.isChecked():
            artist_draw_plot = artist_plot_chart.draw_chart(current_project,seq,'task.status')
        elif self.radbtn_commissioner_status.isChecked():
            artist_draw_plot = artist_plot_chart.draw_chart(current_project,seq,'task.sup_review')
        self.artist_plot_chart_lay.addWidget(artist_draw_plot)
        
    def draw_project_pie(self):
        current_project = self.combox_project.currentText()
        seq = self.get_select_seq()
        
        del_project_lay = self.project_pie_chart_lay.takeAt(0)
        del_project_lay.widget().deleteLater()
        if self.radbtn_task_status.isChecked():
            project_draw_pie = project_pie_chart.draw_chart(current_project,seq,'task.status')
        elif self.radbtn_commissioner_status.isChecked():
            project_draw_pie = project_pie_chart.draw_chart(current_project,seq,'task.sup_review')
        self.project_pie_chart_lay.addWidget(project_draw_pie)
                
    def show_dashboard(self):
        self.draw_pipeline_chart()
        self.draw_artist_chart()
        self.draw_project_pie()
    
if __name__ == "__main__":
    app = QtWidgets.QApplication()
    window = dashboard()
    window.show()
    app.exec_()