# -*- coding: utf-8 -*-
from PySide2.QtCore import *
from PySide2.QtWidgets import *
from PySide2 import QtCore, QtWidgets ,QtGui

import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib import font_manager as fm
import numpy as np

import cgt_info as cgt

class draw_chart(FigureCanvas):
    
    def __init__(self,project=None,seq=None,status_type=None):
        super().__init__()
        
        self.project = project
        self.sequence = seq
        self.status_type = status_type
        
        self.figure = Figure()
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.axis('off')
        
        if project != None:
            self.figure.clear()
            self.canvas = FigureCanvas(self.figure)
            self.pie_chart()
        else:
            ax = self.figure.subplots()
            ax.axis('off')
            ax.set_title('Project Total')
            self.canvas = FigureCanvas(self.figure)
        
    def get_project_data(self):
        proj = cgt.project.filter_project(self.project)[0]
        if self.sequence == None:
            project = cgt.shot.get_project_count(proj,status_type=self.status_type)
        else:
            project = cgt.shot.get_project_count(proj,self.sequence,status_type=self.status_type)
        
        status_color = cgt.project.get_status_color()
        
        status = []
        count = []
        status_color_list = []
        for value_item in project:
            for k_value,v_value in value_item.items():
                status.append(k_value)
                count.append(v_value)
                status_color_list.append(str(status_color[k_value]))
                
        return status,count,status_color_list
        
    def pie_chart(self):
        self.figure.clear()
        
        ax = self.figure.subplots()
        ax.axis('off')
        # status_color = cgt.project.get_status_color()
        data = self.get_project_data()
        
        status = data[0]
        value = data[1]
        status_color = data[2]
        
        explode = []
        for i in range(len(status)):
            if float(value[i]) < 1:
                explode.append(0.15)
            elif float(value[i]) < 2 and float(value[i]) > 1:
                explode.append(0.075)
            else:
                explode.append(0)
            
        # explode = tuple([0 for i in status])
        
        label_size = fm.FontProperties()
        label_size.set_size('small')
        value_size = fm.FontProperties()
        value_size.set_size('x-small')
        
        ax = self.figure.subplots()
        patches, texts, autotexts = ax.pie(value, explode=tuple(explode) ,labels=status, autopct='%1.2f%%',colors=status_color,startangle=90,labeldistance=1.2,pctdistance=1.05)
        
        ax.legend(patches,status,ncols=1,loc=(1.1,0.5),fontsize=6)
        
        plt.setp(autotexts, fontproperties=value_size)
        plt.setp(texts, fontproperties=label_size)
        # patches, texts, autotexts = ax.pie(value, labels=status, autopct='%1.2f%%',explode=explode,
        # shadow=False, startangle=170, colors=status_color, labeldistance=1.2,pctdistance=1.03, radius=0.4)
        
        self.canvas.draw()

# if __name__ == '__main__':
#     app = QtWidgets.QApplication()
#     window = draw_chart('YMDZX',seq=['VFX'],status_type='task.status')
#     # window = draw_chart()
#     window.show()
#     app.exec_()