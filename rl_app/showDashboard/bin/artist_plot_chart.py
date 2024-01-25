# -*- coding: utf-8 -*-
from PySide2.QtCore import *
from PySide2.QtWidgets import *
from PySide2 import QtCore, QtWidgets ,QtGui

import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
import numpy as np

import cgt_info as cgt

class draw_chart(FigureCanvas):
    
    # def __init__(self,project=None,seq=None):
    def __init__(self,project=None,seq=None,status_type=None):
        super().__init__()  
        
        self.project = project
        self.sequence = seq
        self.status_type = status_type
        self.figure = Figure()
        plt.rcParams['font.sans-serif'] = ['SimHei']
        
        if project != None:
            self.figure.clear()
            self.canvas = FigureCanvas(self.figure)
            
            self.plot_chart()
        else:
            ax = self.figure.subplots()
            ax.axis('off')
            # ax.spines['top'].set_visible(False)
            # ax.spines['right'].set_visible(False)
            ax.set_title('Artists Total')
            ax.set_ylabel('Task Number')
            self.canvas = FigureCanvas(self.figure)
        # self.canvas = FigureCanvas(self.figure)
        # self.plot_chart()
        
    def get_artist_data(self):
        artists = cgt.project.get_artist(self.project)
        artists.sort()
        
        proj = cgt.project.filter_project(self.project)[0]
        status = cgt.shot.get_task_status(proj,status=self.status_type)
        status.sort()
        
        init_value = [0 for i in status]

        artist_list = []
        status_value_list = []
        for art in artists:
            if self.sequence == None:
                value = cgt.shot.get_task_artist_count(proj,artist=art,stauts_type=self.status_type)
            else:
                value = cgt.shot.get_task_artist_count(proj,seq=self.sequence,artist=art,stauts_type=self.status_type)
            status_init = dict(zip(status,init_value))
            artist_list.append(art)
            for v_index in value:
                status_init.update(v_index)
            # artsit_data[art] = status_init
            status_value_list.append(status_init)
        
        return artist_list,status_value_list
        
    def plot_chart(self):
        artist_datas = self.get_artist_data()
        self.figure.clear()
        
        ax = self.figure.subplots()
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        status_color = cgt.project.get_status_color()
        
        artists = artist_datas[0]
        artist_value = artist_datas[1]
        
        status_value = {}
        for art_value in artist_value:
            for k_status,v_status in art_value.items():
                status_value.setdefault(k_status,[]).append(int(v_status))
        
        x_artist = tuple(artist_datas[0])
        y_count = status_value
        bottom = np.zeros(len(x_artist))
        
        for y_label,y_value in y_count.items():
            p = ax.bar(x_artist,y_value,label=y_label,bottom=bottom,color=str(status_color[y_label]))
            bottom += y_value
            for bound_value in p:
                x_a,y_a,width,height = bound_value.get_bbox().bounds
                if str(int(height)) != '0':
                    ax.annotate(str(int(height)),(x_a+width/2,height/2+y_a+width/2-0.5),ha='center',va='center',color='#ffffff')
                else:
                    pass
        ax.legend(ncols=1,loc=(0.99,0.5),fontsize=6)
        
        x_label = ax.get_xticklabels()
        for tick in x_label:
            tick.set_fontsize(8)
        
        ax.set_xticklabels(artist_datas[0],rotation=75)
        ax.set_title('Artists Total')
        ax.set_ylabel('Task Number')
        self.canvas.draw()

# if __name__ == '__main__':
#     app = QtWidgets.QApplication()
#     window = draw_chart('YMDZX',seq=['VFX'],status_type='task.sup_review')
#     window.show()
#     app.exec_()
        