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

    def __init__(self,project=None,seq=None,status_type=None):
        super().__init__()  
        
        self.project = project
        self.sequence = seq
        self.status_type = status_type
        self.figure = Figure()
        plt.rcParams['font.sans-serif'] = ['SimHei']
        
        if project != None:
            # self.figure = Figure()
            self.figure.clear()
            self.canvas = FigureCanvas(self.figure)
            # plt.rcParams['font.sans-serif'] = ['SimHei']
            
            self.plot_chart()
        else:
            # self.figure = Figure()
            ax = self.figure.subplots()
            ax.axis('off')
            # ax.spines['top'].set_visible(False)
            # ax.spines['right'].set_visible(False)
            ax.set_title('Pipeline Total')
            # ax.set_xlabel('Pipeline')
            ax.set_ylabel('Task Number')
            self.canvas = FigureCanvas(self.figure)
        
    def get_pipeline_data(self):
        proj = cgt.project.filter_project(self.project)[0]
        
        pipeline = cgt.shot.get_task_pipeline(proj)
        pipeline.sort()
        status = cgt.shot.get_task_status(proj,status=self.status_type)
        status.sort()
        
        init_value = [0 for i in status]
        pipeline_list = []
        status_value_list = []
        for pipe in pipeline:
            if self.sequence == None:
                value = cgt.shot.get_task_pipeline_count(proj,pipeline=pipe,stauts_type=self.status_type)
            else:
                value = cgt.shot.get_task_pipeline_count(proj,seq=self.sequence,pipeline=pipe,stauts_type=self.status_type)
            status_init = dict(zip(status,init_value))
            pipeline_list.append(pipe)
            for v_index in value:
                status_init.update(v_index)
            status_value_list.append(status_init)
            
        return pipeline_list,status_value_list
    
    def plot_chart(self):
        self.figure.clear()
        
        ax = self.figure.subplots()
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        
        status_color = cgt.project.get_status_color()
        data = self.get_pipeline_data()
        
        status_value = {}
        for pipe_value in data[1]:
            for k_status,v_status in pipe_value.items():
                status_value.setdefault(k_status,[]).append(int(v_status))

        x_name = tuple(data[0])
        y_count = status_value
        
        bottom = np.zeros(len(x_name))
        
        for y_label , y_value in y_count.items():
            p = ax.bar(x_name,y_value,label=y_label,bottom=bottom,color=str(status_color[y_label]))
            bottom += y_value
            for bound_value in p:
                x_a,y_a,width,height = bound_value.get_bbox().bounds
                if str(int(height)) != '0':
                    ax.annotate(str(int(height)),(x_a+width/2,height/2+y_a+width/2-0.5),ha='center',va='center',color='#ffffff')
                    # ax.bar_label(p, label_type='center',color='#ffffff')
                else:
                    pass
        ax.legend(ncols=1,loc=(1,0.5),fontsize=6)
    
        ax.set_title('Pipeline Total')
        # ax.set_xlabel('Pipeline')
        ax.set_ylabel('Task Number')
        self.canvas.draw()

# if __name__ == '__main__':
#     app = QtWidgets.QApplication()
#     window = draw_chart('YMDZX',['VFX'])
#     window.show()
#     app.exec_()
