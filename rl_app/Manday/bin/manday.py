# -*- coding: utf-8 -*-

import os
import sys
from PySide2.QtCore import *
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from PySide2.QtGui import QMouseEvent ,Qt ,QCursor
from PySide2 import QtCore, QtWidgets ,QtGui

import xlrd
import xlwt
import json
import datetime
from dateutil.relativedelta import relativedelta

current_dir = os.path.dirname(__file__)
root_dir = os.path.dirname(os.path.dirname(current_dir))
sys.path.append(root_dir)
sys.path.append(current_dir)
import cgt_info

class Man_day(QMainWindow):
    
    excel_Signal = QtCore.Signal(list)
    
    def setupUi(self,MainWindow):
        MainWindow.setObjectName(u"Man-day")
        MainWindow.resize(1480,800)
        MainWindow.setWindowTitle(u"人天工具")
        # MainWindow.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        # MainWindow.setFixedSize(MainWindow.width(),MainWindow.height())
        
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.show_lay = QGridLayout(MainWindow)
        self.show_lay.setObjectName(u"人天工具")
        
        self.lab_projects = QLabel('Projects')
        self.combox_project = custom_comboBox()
        projects = cgt_info.info.get_active_projects()
        projects.sort()
        self.combox_project.add_items(projects)
        
        self.lab_year = QLabel('Year')
        self.comb_year = QComboBox()
        current_year = datetime.datetime.now().year
        years = [str(current_year-1),str(current_year),str(current_year+1)]
        self.comb_year.addItems(years)
        self.comb_year.setCurrentIndex(1)
        
        self.checkbox_only = QPushButton('Select Only')
        self.checkbox_only.setCheckable(1)
        
        self.radbtn_artist = QRadioButton('Artist')
        self.radbtn_project = QRadioButton('Project')
        self.radbtn_artist.setChecked(1)
        self.radbtn_project.setChecked(0)
        
        self.lab_filter = QLabel('User:')
        current_file = os.path.abspath(__file__)
        self.json_root = os.path.dirname(current_file)

        self.list_user = QTreeWidget()
        # self.list_user.setSelectionMode(QTreeWidget.NoSelection)
        self.list_user.setHeaderHidden(1)
        self.get_user_list()
            
        ########## combox_artists set ##########
        
        self.btn_setting = QPushButton('Artist Setting')
        self.btn_refresh = QPushButton('Refresh')
        self.btn_export = QPushButton('Export')
        
        self.main_table = QTreeWidget()
        self.main_table.setAlternatingRowColors(True)
        main_table_policy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)
        main_table_policy.setHorizontalStretch(100)
        main_table_policy.setHeightForWidth(self.main_table.sizePolicy().hasHeightForWidth())
        self.main_table.setSizePolicy(main_table_policy)
        self.main_table.setMinimumSize(QSize(1150,0))
        
        self.main_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.main_table.header().setDefaultAlignment(QtCore.Qt.AlignCenter)
                
        self.generate_main_by_artist()
        
        self.form_projects = QFormLayout()
        self.form_projects.setWidget(0,QFormLayout.LabelRole,self.lab_projects)
        self.form_projects.setWidget(0,QFormLayout.FieldRole,self.combox_project)
        
        self.form_year = QFormLayout()
        self.form_year.setWidget(0,QFormLayout.LabelRole,self.lab_year)
        self.form_year.setWidget(0,QFormLayout.FieldRole,self.comb_year)
        
        self.set_lay = QHBoxLayout()
        self.set_lay.addWidget(self.btn_setting)
        self.set_lay.addWidget(self.btn_refresh)
        
        self.common_lay = QHBoxLayout()
        self.common_lay.addLayout(self.form_projects)
        self.common_lay.addLayout(self.form_year)
        
        self.radbtn_lay = QHBoxLayout()
        self.radbtn_lay.addWidget(self.radbtn_artist)
        self.radbtn_lay.addWidget(self.radbtn_project)
        
        self.ver_userlay = QVBoxLayout()
        self.ver_userlay.addWidget(self.checkbox_only)
        self.ver_userlay.addLayout(self.radbtn_lay)
        self.ver_userlay.addWidget(self.list_user)
        
        self.main_lay = QHBoxLayout()
        self.main_lay.addLayout(self.ver_userlay)
        self.main_lay.addWidget(self.main_table)
        
        self.show_lay.addLayout(self.common_lay,0,0,1,3)
        self.show_lay.addLayout(self.set_lay,0,4,1,1)
        self.show_lay.addLayout(self.main_lay,1,0,1,5)
        self.show_lay.addWidget(self.btn_export,2,4,1,1)
        
        self.combox_project.currentTextChanged.connect(self.fliter_project)
        self.main_table.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.main_table.customContextMenuRequested.connect(self.showContextMenu_maintable)

        self.list_user.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.list_user.customContextMenuRequested.connect(self.showContextMenu_user)
        self.list_user.clicked.connect(self.check_user_stage,Qt.UniqueConnection)

        self.btn_export.clicked.connect(self.show_export)
        self.btn_refresh.clicked.connect(self.refresh)
        self.btn_setting.clicked.connect(self.setting)
        self.checkbox_only.clicked.connect(self.checkbox_stage,Qt.UniqueConnection)
        self.radbtn_artist.clicked.connect(self.change_artist_project)
        self.radbtn_project.clicked.connect(self.change_artist_project)
        self.comb_year.currentTextChanged.connect(self.change_artist_project)
        
    def setting(self):
        self.setting_window = artist_setting()
        self.setting_window.show()
        
    def get_user_list(self):
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        try:
            json_path = '%s/user.json'%(json_root)
            if os.path.exists(json_path):
                os.chmod(json_path,0o755)
                with open(json_path,'r+',encoding='utf-8') as f:
                    self.users = json.load(f)
            else:
                return 'user.json is not exist!'
        except:
            return 'user.json is not exist!'
        
        for group in self.users:
            group_item = QTreeWidgetItem(self.list_user)
            group_item.setText(0,group)
            group_item.setCheckState(0,Qt.Checked)
            group_item.setExpanded(1)
            group_item.setFlags(Qt.ItemFlag.ItemIsAutoTristate | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            # sort_group = self.sort_dict(self.users[group])
            sort_group = self.users[group]
            for artist in sort_group:
                artist_item = QTreeWidgetItem(group_item)
                artist_item.setText(0,sort_group[artist]['cn_name'])
                if sort_group[artist]['check'] == 0:
                    artist_item.setDisabled(1)
                    artist_item.setCheckState(0,Qt.Unchecked)
                else:
                    artist_item.setDisabled(0)
                    artist_item.setCheckState(0,Qt.Checked)

    def change_artist_project(self):
        self.list_user.clear()
        self.main_table.setColumnCount(0)
        self.main_table.clear()
        current_year = self.comb_year.currentText()
        
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        if self.radbtn_artist.isChecked():
            # try:
            #     json_path = '%s/temp_cgtdata.json'%(json_root)
            #     tasks = cgt_info.info.get_tasks_manday()
            #     with open(json_path,'w+',encoding='utf-8') as f:
            #         json.dump(tasks,f,indent=4,ensure_ascii=False)
            # except:
            #     print('Error')
            self.checkbox_only.setEnabled(1)
            self.btn_export.setEnabled(1)
            self.get_user_list()
            self.generate_main_by_artist()
        if self.radbtn_project.isChecked():
            self.checkbox_only.setEnabled(0)
            self.btn_export.setEnabled(0)
            self.get_project_list()
            self.list_user.clicked.connect(self.genera_main_by_project)
            # self.genera_main_by_project()

    def get_project_list(self):
        active_project = cgt_info.info.get_active_projects()
        active_project.sort()
        for index in active_project:
            project_item = QTreeWidgetItem(self.list_user)
            project_item.setText(0,index)
            
    def generate_main_by_artist(self):
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        
        current_year = self.comb_year.currentText()
        try:
            json_path = '%s/temp_cgtdata.json'%(json_root)
            os.chmod(json_path,0o755)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.tasks = json.load(f)
            # if os.path.exists(json_path):
            #     os.chmod(json_path,0o755)
            # else:
            #     tasks = cgt_info.info.get_tasks_manday()
            #     with open(json_path,'w+',encoding='utf-8') as f:
            #         json.dump(tasks,f,indent=4,ensure_ascii=False)
            #     os.chmod(json_path,0o755)
            #     with open(json_path,'r+',encoding='utf-8') as f:
            #         self.tasks = json.load(f)
            #     print('################')
        except:
            json_path = '%s/temp_cgtdata.json'%(json_root)
            tasks = cgt_info.info.get_tasks_manday()
            with open(json_path,'w') as f:
                json.dump(tasks,f)
            # # os.chmod(json_path,0o755)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.tasks = json.load(f)
            self.tasks = cgt_info.info.get_tasks_manday()
                            
        self.header = ['Artists','Total','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sept','Oct','Nov','Dec']
        self.main_table.setHeaderLabels(self.header)
        
        for header_index in range(len(self.header)):
            if header_index == 0:
                self.main_table.setColumnWidth(header_index,200)
            else:
                self.main_table.setColumnWidth(header_index,75)
        
        user_list = []
        user_it = QTreeWidgetItemIterator(self.list_user)
        while user_it.value():
            if user_it.value().isDisabled() == False and user_it.value().parent() != None:
                user_list.append(user_it.value().text(0))
            user_it += 1
            
        task_manday_list = []
        seq_manday_list = []
        task_jan_list = []
        seq_jan_list = []
        task_feb_list = []
        seq_feb_list = []
        task_mar_list = []
        seq_mar_list = []
        task_apr_list = []
        seq_apr_list = []
        task_may_list = []
        seq_may_list = []
        task_jun_list = []
        seq_jun_list = []
        task_jul_list = []
        seq_jul_list = []
        task_aug_list = []
        seq_aug_list = []
        task_sep_list = []
        seq_sep_list = []
        task_oct_list = []
        seq_oct_list = []
        task_nov_list = []
        seq_nov_list = []
        task_dec_list = []
        seq_dec_list = []
            
        for artist in self.tasks:
            if artist in user_list:
                artist_item = QTreeWidgetItem(self.main_table)
                artist_item.setTextAlignment(1,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
                artist_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
                artist_item.setExpanded(False)
                artist_item.setText(0,artist)
                for project in self.tasks[artist]:
                    project_item = QTreeWidgetItem(artist_item)
                    project_item.setTextAlignment(1,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
                    project_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
                    project_item.setExpanded(False)
                    project_item.setText(0,project)
                    for seq in self.tasks[artist][project]:
                        seq_item = QTreeWidgetItem(project_item)
                        seq_item.setTextAlignment(1,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
                        seq_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
                        seq_item.setExpanded(False)
                        seq_item.setText(0,seq)
                        for task in self.tasks[artist][project][seq]:
                            update_year = datetime.datetime.strptime(self.tasks[artist][project][seq][task]['update_time'],'%Y-%m-%d %H:%M:%S')
                            create_time = datetime.datetime.strptime(self.tasks[artist][project][seq][task]['create_time'],'%Y-%m-%d %H:%M:%S')
                            work_year = self.tasks[artist][project][seq][task]['project_workyear']
                            task_item = QTreeWidgetItem(seq_item)
                            task_item.setTextAlignment(1,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
                            task_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
                            task_item.setExpanded(False)
                            # creat_time = self.tasks[artist][project][seq][task]['create_time']
                            # update_time = datetime.datetime.strptime(self.tasks[artist][project][seq][task]['update_time'],'%Y-%m-%d %H:%M:%S')
                            # compare_time = update_time + relativedelta(months=1)
                            # if str(compare_time.year) == current_year:
                            #     print(compare_time.year)
                            task_item.setText(0,task)
                            task_item.setText(1,self.tasks[artist][project][seq][task]['mandays'])
                            try:
                                is_total_num = float(self.tasks[artist][project][seq][task]['mandays'])
                                task_manday_list.append(is_total_num)
                            except:
                                pass
                            task_item.setText(2,self.tasks[artist][project][seq][task]['jan'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_jan_num = float(self.tasks[artist][project][seq][task]['jan'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_jan_num = float(self.tasks[artist][project][seq][task]['jan'])
                                            # task_jan_list.append(is_jan_num)
                                        else:
                                            is_jan_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_jan_num = float(self.tasks[artist][project][seq][task]['jan'])
                                    else:
                                        is_jan_num = float('0')
                                task_jan_list.append(is_jan_num)
                            except:
                                pass
                            task_item.setText(3,self.tasks[artist][project][seq][task]['feb'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_feb_num = float(self.tasks[artist][project][seq][task]['feb'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_feb_num = float(self.tasks[artist][project][seq][task]['feb'])
                                            # task_feb_list.append(is_feb_num)
                                        else:
                                            is_feb_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_feb_num = float(self.tasks[artist][project][seq][task]['feb'])
                                    else:
                                        is_feb_num = float('0')
                                task_feb_list.append(is_feb_num)
                            except:
                                pass
                            task_item.setText(4,self.tasks[artist][project][seq][task]['mar'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_mar_num = float(self.tasks[artist][project][seq][task]['mar'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_mar_num = float(self.tasks[artist][project][seq][task]['mar'])
                                            # task_mar_list.append(is_mar_num)
                                        else:
                                            is_mar_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_mar_num = float(self.tasks[artist][project][seq][task]['mar'])
                                    else:
                                        is_mar_num = float('0')
                                task_mar_list.append(is_mar_num)
                            except:
                                pass
                            task_item.setText(5,self.tasks[artist][project][seq][task]['apr'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_apr_num = float(self.tasks[artist][project][seq][task]['apr'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_apr_num = float(self.tasks[artist][project][seq][task]['apr'])
                                            # task_apr_list.append(is_apr_num)
                                        else:
                                            is_apr_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_apr_num = float(self.tasks[artist][project][seq][task]['apr'])
                                    else:
                                        is_apr_num = float('0')
                                task_apr_list.append(is_apr_num)
                            except:
                                pass
                            task_item.setText(6,self.tasks[artist][project][seq][task]['may'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_may_num = float(self.tasks[artist][project][seq][task]['may'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_may_num = float(self.tasks[artist][project][seq][task]['may'])
                                            # task_may_list.append(is_may_num)
                                        else:
                                            is_may_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_may_num = float(self.tasks[artist][project][seq][task]['may'])
                                    else:
                                        is_may_num = float('0')
                                task_may_list.append(is_may_num)
                            except:
                                pass
                            task_item.setText(7,self.tasks[artist][project][seq][task]['jun'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_jun_num = float(self.tasks[artist][project][seq][task]['jun'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_jun_num = float(self.tasks[artist][project][seq][task]['jun'])
                                            # task_jun_list.append(is_jun_num)
                                        else:
                                            is_jun_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_jun_num = float(self.tasks[artist][project][seq][task]['jun'])
                                    else:
                                        is_jun_num = float('0')
                                task_jun_list.append(is_jun_num)
                            except:
                                pass
                            task_item.setText(8,self.tasks[artist][project][seq][task]['jul'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_jul_num = float(self.tasks[artist][project][seq][task]['jul'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_jul_num = float(self.tasks[artist][project][seq][task]['jul'])
                                            # task_jul_list.append(is_jul_num)
                                        else:
                                            is_jul_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_jul_num = float(self.tasks[artist][project][seq][task]['jul'])
                                    else:
                                        is_jul_num = float('0')
                                task_jul_list.append(is_jul_num)
                            except:
                                pass
                            task_item.setText(9,self.tasks[artist][project][seq][task]['aug'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_aug_num = float(self.tasks[artist][project][seq][task]['aug'])
                                else:
                                    if work_year == '':
                                        if str((update_year + relativedelta(months=4)).year) == str(int(current_year)):
                                            is_aug_num = float(self.tasks[artist][project][seq][task]['aug'])
                                            # task_aug_list.append(is_aug_num)
                                        else:
                                            is_aug_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[1]:
                                        is_aug_num = float(self.tasks[artist][project][seq][task]['aug'])
                                    else:
                                        is_aug_num = float('0')
                                task_aug_list.append(is_aug_num)
                            except:
                                pass
                            task_item.setText(10,self.tasks[artist][project][seq][task]['sep'])
                            try:
                                if str(int(current_year)) == work_year:
                                    is_sep_num = float(self.tasks[artist][project][seq][task]['sep'])
                                else:
                                    if work_year == '':
                                        if str(create_time.year) == str(int(current_year)) and str(update_year.year) == str(int(current_year)):
                                            is_sep_num = float(self.tasks[artist][project][seq][task]['sep'])
                                        else:
                                            is_sep_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[0]:
                                        is_sep_num = float(self.tasks[artist][project][seq][task]['sep'])
                                    else:
                                        is_sep_num = float('0')
                                # if str(update_year.month) >= str(12):
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_sep_num = float(self.tasks[artist][project][seq][task]['sep'])
                                #     else:
                                #         is_sep_num = float('0')
                                # else:
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_sep_num = float(self.tasks[artist][project][seq][task]['sep'])
                                #     else:
                                #         is_sep_num = float('0')
                                task_sep_list.append(is_sep_num)
                            except:
                                pass
                            task_item.setText(11,self.tasks[artist][project][seq][task]['oct'])
                            try:
                                # if str(update_year.month) >= str(10):
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_oct_num = float(self.tasks[artist][project][seq][task]['oct'])
                                #     else:
                                #         is_oct_num = float('0')
                                # else:
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_oct_num = float(self.tasks[artist][project][seq][task]['oct'])
                                #     else:
                                #         is_oct_num = float('0')
                                if str(int(current_year)) == work_year:
                                    is_oct_num = float(self.tasks[artist][project][seq][task]['oct'])
                                else:
                                    if work_year == '':
                                        if str(create_time.year) == str(int(current_year)) and str(update_year.year) == str(int(current_year)):
                                            is_oct_num = float(self.tasks[artist][project][seq][task]['oct'])
                                        else:
                                            is_oct_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[0]:
                                        is_oct_num = float(self.tasks[artist][project][seq][task]['oct'])
                                    else:
                                        is_oct_num = float('0')
                                task_oct_list.append(is_oct_num)
                            except:
                                pass
                            task_item.setText(12,self.tasks[artist][project][seq][task]['nov'])
                            try:
                                # if str(update_year.month) >= str(11):
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_nov_num = float(self.tasks[artist][project][seq][task]['nov'])
                                #     else:
                                #         is_nov_num = float('0')
                                # else:
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_nov_num = float(self.tasks[artist][project][seq][task]['nov'])
                                #     else:
                                #         is_nov_num = float('0')
                                
                                if str(int(current_year)) == work_year:
                                    is_nov_num = float(self.tasks[artist][project][seq][task]['nov'])
                                else:
                                    if work_year == '':
                                        if str(create_time.year) == str(int(current_year)) and str(update_year.year) == str(int(current_year)):
                                            is_nov_num = float(self.tasks[artist][project][seq][task]['nov'])
                                        else:
                                            is_nov_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[0]:
                                        is_nov_num = float(self.tasks[artist][project][seq][task]['nov'])
                                    else:
                                        is_nov_num = float('0')
                                task_nov_list.append(is_nov_num)
                            except:
                                pass
                            task_item.setText(13,self.tasks[artist][project][seq][task]['dec'])
                            try:
                                # if str(create_time.year) == str(int(current_year)):
                                #     is_dec_num = float(self.tasks[artist][project][seq][task]['dec'])
                                # else:
                                #     is_dec_num = float('0')
                                
                                if str(int(current_year)) == work_year:
                                    is_dec_num = float(self.tasks[artist][project][seq][task]['dec'])
                                else:
                                    if work_year == '':
                                        if str(create_time.year) == str(int(current_year)) and str(update_year.year) == str(int(current_year)):
                                            is_dec_num = float(self.tasks[artist][project][seq][task]['dec'])
                                        else:
                                            is_dec_num = float('0')
                                    elif str(int(current_year)) == work_year.split('-')[0]:
                                        is_dec_num = float(self.tasks[artist][project][seq][task]['dec'])
                                    else:
                                        is_dec_num = float('0')
                                # if str(update_year.month) >= str(12):
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_dec_num = float(self.tasks[artist][project][seq][task]['dec'])
                                #     else:
                                #         is_dec_num = float('0')
                                # else:
                                #     if str(update_year.year) == str(int(current_year)):
                                #         is_dec_num = float(self.tasks[artist][project][seq][task]['dec'])
                                #     else:
                                #         is_dec_num = float('0')
                                task_dec_list.append(is_dec_num)
                            except:
                                pass
                        task_total_num = "%.1f"%(sum(task_manday_list))
                        if task_total_num != '0.0':
                            seq_item.setText(1,task_total_num)
                        seq_manday_list.append(sum(task_manday_list))
                        task_manday_list.clear()
                        task_jan_num = "%.1f"%(sum(task_jan_list))
                        if task_jan_num != '0.0':
                            seq_item.setText(2,task_jan_num)
                        seq_jan_list.append(sum(task_jan_list))
                        task_jan_list.clear()
                        task_feb_num = "%.1f"%(sum(task_feb_list))
                        if task_feb_num != '0.0':
                            seq_item.setText(3,task_feb_num)
                        seq_feb_list.append(sum(task_feb_list))
                        task_feb_list.clear()
                        task_mar_num = "%.1f"%(sum(task_mar_list))
                        if task_mar_num != '0.0':
                            seq_item.setText(4,task_mar_num)
                        seq_mar_list.append(sum(task_mar_list))
                        task_mar_list.clear()
                        task_apr_num = "%.1f"%(sum(task_apr_list))
                        if task_apr_num != '0.0':
                            seq_item.setText(5,task_may_num)
                        seq_apr_list.append(sum(task_apr_list))
                        task_apr_list.clear()
                        task_may_num = "%.1f"%(sum(task_may_list))
                        if task_may_num != '0.0':
                            seq_item.setText(6,task_may_num)
                        seq_may_list.append(sum(task_may_list))
                        task_may_list.clear()
                        task_jun_num = "%.1f"%(sum(task_jun_list))
                        if task_jun_num != '0.0':
                            seq_item.setText(7,task_jun_num)
                        seq_jun_list.append(sum(task_jun_list))
                        task_jun_list.clear()
                        task_jul_num = "%.1f"%(sum(task_jul_list))
                        if task_jul_num != '0.0':
                            seq_item.setText(8,task_jul_num)
                        seq_jul_list.append(sum(task_jul_list))
                        task_jul_list.clear()
                        task_aug_num = "%.1f"%(sum(task_aug_list))
                        if task_aug_num != '0.0':
                            seq_item.setText(9,task_aug_num)
                        seq_aug_list.append(sum(task_aug_list))
                        task_aug_list.clear()
                        task_sep_num = "%.1f"%(sum(task_sep_list))
                        if task_sep_num != '0.0':
                            seq_item.setText(10,task_sep_num)
                        seq_sep_list.append(sum(task_sep_list))
                        task_sep_list.clear()
                        task_oct_num = "%.1f"%(sum(task_oct_list))
                        if task_oct_num != '0.0':
                            seq_item.setText(11,task_oct_num)
                        seq_oct_list.append(sum(task_oct_list))
                        task_oct_list.clear()
                        task_nov_num = "%.1f"%(sum(task_nov_list))
                        if task_nov_num != '0.0':
                            seq_item.setText(12,task_nov_num)
                        seq_nov_list.append(sum(task_nov_list))
                        task_nov_list.clear()
                        task_dec_num = "%.1f"%(float(sum(task_dec_list)))
                        if task_dec_num != '0.0':
                            seq_item.setText(13,task_dec_num)
                        seq_dec_list.append(sum(task_dec_list))
                        task_dec_list.clear()
                    
                    project_total_num = "%.1f"%(sum(seq_manday_list))
                    if project_total_num != '0.0':
                        project_item.setText(1,project_total_num)
                    seq_manday_list.clear()
                    project_jan_num = "%.1f"%(sum(seq_jan_list))
                    if project_jan_num != '0.0':
                        project_item.setText(2,project_jan_num)
                    seq_jan_list.clear()
                    project_feb_num = "%.1f"%(sum(seq_feb_list))
                    if project_feb_num != '0.0':
                        project_item.setText(3,project_feb_num)
                    seq_feb_list.clear()
                    project_mar_num = "%.1f"%(sum(seq_mar_list))
                    if project_mar_num != '0.0':
                        project_item.setText(4,project_mar_num)
                    seq_mar_list.clear()
                    project_apr_num = "%.1f"%(sum(seq_apr_list))
                    if project_apr_num != '0.0':
                        project_item.setText(5,project_apr_num)
                    seq_apr_list.clear()
                    project_may_num = "%.1f"%(sum(seq_may_list))
                    if project_may_num != '0.0':
                        project_item.setText(6,project_may_num)
                    seq_may_list.clear()
                    project_jun_num = "%.1f"%(sum(seq_jun_list))
                    if project_jun_num != '0.0':
                        project_item.setText(7,project_jun_num)
                    seq_jun_list.clear()
                    project_jul_num = "%.1f"%(sum(seq_jul_list))
                    if project_jul_num != '0.0':
                        project_item.setText(8,project_jul_num)
                    seq_jul_list.clear()
                    project_aug_num = "%.1f"%(sum(seq_aug_list))
                    if project_aug_num != '0.0':
                        project_item.setText(9,project_aug_num)
                    seq_aug_list.clear()
                    project_sep_num = "%.1f"%(sum(seq_sep_list))
                    if project_sep_num != '0.0':
                        project_item.setText(10,project_sep_num)
                    seq_sep_list.clear()
                    project_oct_num = "%.1f"%(sum(seq_oct_list))
                    if project_oct_num != '0.0':
                        project_item.setText(11,project_oct_num)
                    seq_oct_list.clear()
                    project_nov_num = "%.1f"%(sum(seq_nov_list))
                    if project_nov_num != '0.0':
                        project_item.setText(12,project_nov_num)
                    seq_nov_list.clear()
                    project_dec_num = "%.1f"%(sum(seq_dec_list))
                    if project_dec_num != '0.0':
                        project_item.setText(13,project_dec_num)
                    seq_dec_list.clear()
                total_amount_item = QTreeWidgetItem(artist_item)
                total_amount_item.setText(0,'Total-Amount')

        artist_count = self.main_table.topLevelItemCount()
        for top_index in range(artist_count):
            project_total_manday = []
            project_jan_manday = []
            project_feb_manday = []
            project_mar_manday = []
            project_apr_manday = []
            project_may_manday = []
            project_jun_manday = []
            project_jul_manday = []
            project_aug_manday = []
            project_sep_manday = []
            project_oct_manday = []
            project_nov_manday = []
            project_dec_manday = []
            top_item = self.main_table.topLevelItem(top_index)
            child_count = top_item.childCount()
            for child_index in range(child_count):
                child_item = top_item.child(child_index)
                child_total_manday = child_item.text(1)
                child_jan_manday = child_item.text(2)
                child_feb_manday = child_item.text(3)
                child_mar_manday = child_item.text(4)
                child_apr_manday = child_item.text(5)
                child_may_manday = child_item.text(6)
                child_jun_manday = child_item.text(7)
                child_jul_manday = child_item.text(8)
                child_aug_manday = child_item.text(9)
                child_sep_manday = child_item.text(10)
                child_oct_manday = child_item.text(11)
                child_nov_manday = child_item.text(12)
                child_dec_manday = child_item.text(13)
                if child_total_manday != '':
                    project_total_manday.append(float(child_total_manday))
                if child_jan_manday != '':
                    project_jan_manday.append(float(child_jan_manday))
                if child_feb_manday != '':
                    project_feb_manday.append(float(child_feb_manday))
                if child_mar_manday != '':
                    project_mar_manday.append(float(child_mar_manday))
                if child_apr_manday != '':
                    project_apr_manday.append(float(child_apr_manday))
                if child_may_manday != '':
                    project_may_manday.append(float(child_may_manday))
                if child_jun_manday != '':
                    project_jun_manday.append(float(child_jun_manday))
                if child_jul_manday != '':
                    project_jul_manday.append(float(child_jul_manday))
                if child_aug_manday != '':
                    project_aug_manday.append(float(child_aug_manday))
                if child_sep_manday != '':
                    project_sep_manday.append(float(child_sep_manday))
                if child_oct_manday != '':
                    project_oct_manday.append(float(child_oct_manday))
                if child_nov_manday != '':
                    project_nov_manday.append(float(child_nov_manday))
                if child_dec_manday != '':
                    project_dec_manday.append(float(child_dec_manday))
            if sum(project_total_manday) != 0.0:
                top_item.setText(1,"%.1f"%(sum(project_total_manday)))
            else:
                top_item.setText(1,'')
            if sum(project_jan_manday) != 0.0:
                top_item.setText(2,"%.1f"%(sum(project_jan_manday)))
                if sum(project_jan_manday) >= 24 and sum(project_dec_manday) < 40:
                    top_item.setBackgroundColor(2,QColor(255,159,24,255))
                if sum(project_jan_manday) >= 40:
                    top_item.setBackgroundColor(2,QColor(190,0,0,255))
                    top_item.setForeground(2,QColor(255,255,255,255))
            else:
                top_item.setText(2,'')
            if sum(project_feb_manday) != 0.0:
                top_item.setText(3,"%.1f"%(sum(project_feb_manday)))
                if sum(project_feb_manday) >= 24:
                    top_item.setBackgroundColor(3,QColor(255,159,24,255))
            else:
                top_item.setText(3,'')
            if sum(project_mar_manday) != 0.0:
                top_item.setText(4,"%.1f"%(sum(project_mar_manday)))
                if sum(project_mar_manday) >= 24:
                    top_item.setBackgroundColor(4,QColor(255,159,24,255))
            else:
                top_item.setText(4,'')
            if sum(project_apr_manday) != 0.0:
                top_item.setText(5,"%.1f"%(sum(project_apr_manday)))
                if sum(project_apr_manday) >= 24:
                    top_item.setBackgroundColor(5,QColor(255,159,24,255))
            else:
                top_item.setText(5,'')
            if sum(project_may_manday) != 0.0:
                top_item.setText(6,"%.1f"%(sum(project_may_manday)))
                if sum(project_may_manday) >= 24:
                    top_item.setBackgroundColor(6,QColor(255,159,24,255))
            else:
                top_item.setText(6,'')
            if sum(project_jun_manday) != 0.0:
                top_item.setText(7,"%.1f"%(sum(project_jun_manday)))
                if sum(project_jun_manday) >= 24:
                    top_item.setBackgroundColor(7,QColor(255,159,24,255))
            else:
                top_item.setText(7,'')
            if sum(project_jul_manday) != 0.0:
                top_item.setText(8,"%.1f"%(sum(project_jul_manday)))
                if sum(project_jul_manday) >= 24:
                    top_item.setBackgroundColor(8,QColor(255,159,24,255))
            else:
                top_item.setText(8,'')
            if sum(project_aug_manday) != 0.0:
                top_item.setText(9,"%.1f"%(sum(project_aug_manday)))
                if sum(project_aug_manday) >= 24:
                    top_item.setBackgroundColor(9,QColor(255,159,24,255))
            else:
                top_item.setText(9,'')
            if sum(project_sep_manday) != 0.0:
                top_item.setText(10,"%.1f"%(sum(project_sep_manday)))
                if sum(project_sep_manday) >= 24:
                    top_item.setBackgroundColor(10,QColor(255,159,24,255))
            else:
                top_item.setText(10,'')
            if sum(project_oct_manday) != 0.0:
                top_item.setText(11,"%.1f"%(sum(project_oct_manday)))
                if sum(project_oct_manday) >= 24:
                    top_item.setBackgroundColor(11,QColor(255,159,24,255))
            else:
                top_item.setText(11,'')
            if sum(project_nov_manday) != 0.0:
                top_item.setText(12,"%.1f"%(sum(project_nov_manday)))
                if sum(project_nov_manday) >= 24:
                    top_item.setBackgroundColor(12,QColor(255,159,24,255))
            else:
                top_item.setText(12,'')
            if sum(project_dec_manday) != 0.0:
                top_item.setText(13,"%.1f"%(sum(project_dec_manday)))
                if sum(project_dec_manday) >= 24 and sum(project_dec_manday) < 40:
                    top_item.setBackgroundColor(13,QColor(255,159,24,255))
                if sum(project_dec_manday) >= 40:
                    top_item.setBackgroundColor(13,QColor(255,0,0,200))
                    top_item.setForeground(13,QColor(255,255,255,255))
            else:
                top_item.setText(13,'')
                
        self.genera_total_item()

        main_it = QTreeWidgetItemIterator(self.main_table)
        while main_it.value():
            for col in range(len(self.header)):
                item_frame = QFrame()
                item_frame.setFrameShape(QFrame.Box)
                item_frame.setStyleSheet("QFrame{border:0.5px solid #E5E5E5};")
                self.main_table.setItemWidget(main_it.value(),col,item_frame)
            main_it += 1
        
    def sort_dict(self,dict_info):
        reslut = {}
        for i in sorted([i for i in dict_info.keys()]):
            v = dict_info.get(i)
            if isinstance(v,dict):
                v = self.sort_dict(v)
            reslut[i] = v
            
        return reslut
    
    def genera_total_item(self):
        total_manday = []
        total_jan = []
        total_feb = []
        total_mar = []
        total_apr = []
        total_may = []
        total_jun = []
        total_jul = []
        total_aug = []
        total_sep = []
        total_oct = []
        total_nov = []
        total_dec = []
        
        top_count = self.main_table.topLevelItemCount()
        for top_index in range(top_count):
            top_item = self.main_table.topLevelItem(top_index)
            manday = top_item.text(1)
            jan = top_item.text(2)
            feb = top_item.text(3)
            mar = top_item.text(4)
            apr = top_item.text(5)
            may = top_item.text(6)
            jun = top_item.text(7)
            jul = top_item.text(8)
            aug = top_item.text(9)
            sep = top_item.text(10)
            oct = top_item.text(11)
            nov = top_item.text(12)
            dec = top_item.text(13)
            if manday != '':
                total_manday.append(float(manday))
            if jan != '':
                total_jan.append(float(jan))
            if feb != '':
                total_feb.append(float(feb))
            if mar != '':
                total_mar.append(float(mar))
            if apr != '':
                total_apr.append(float(apr))
            if may != '':
                total_may.append(float(may))
            if jun != '':
                total_jun.append(float(jun))
            if jul != '':
                total_jul.append(float(jul))
            if aug != '':
                total_aug.append(float(aug))
            if sep != '':
                total_sep.append(float(sep))
            if oct != '':
                total_oct.append(float(oct))
            if nov != '':
                total_nov.append(float(nov))
            if dec != '':
                total_dec.append(float(dec))
        
        self.all_sum_item = QTreeWidgetItem(self.main_table)
        self.all_sum_item.setText(0,'Sum Man-day')
        self.all_sum_item.setText(1,"%.1f"%(sum(total_manday)))
        self.all_sum_item.setText(2,"%.1f"%(sum(total_jan)))
        self.all_sum_item.setText(3,"%.1f"%(sum(total_feb)))
        self.all_sum_item.setText(4,"%.1f"%(sum(total_mar)))
        self.all_sum_item.setText(5,"%.1f"%(sum(total_apr)))
        self.all_sum_item.setText(6,"%.1f"%(sum(total_may)))
        self.all_sum_item.setText(7,"%.1f"%(sum(total_jun)))
        self.all_sum_item.setText(8,"%.1f"%(sum(total_jul)))
        self.all_sum_item.setText(9,"%.1f"%(sum(total_aug)))
        self.all_sum_item.setText(10,"%.1f"%(sum(total_sep)))
        self.all_sum_item.setText(11,"%.1f"%(sum(total_oct)))
        self.all_sum_item.setText(12,"%.1f"%(sum(total_nov)))
        self.all_sum_item.setText(13,"%.1f"%(sum(total_dec)))
        self.all_sum_item.setTextAlignment(1,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
        
        total_dec.clear()

    def genera_total_item_by_project(self):
        total_manday = []
        total_input = []
        total_dm = []
        total_lay = []
        total_ani = []
        total_cfx = []
        total_fx = []
        total_lig = []
        total_mp = []
        total_roto = []
        total_paint = []
        total_comp = []
        total_output = []
        
        top_count = self.main_table.topLevelItemCount()
        for top_index in range(top_count):
            top_item = self.main_table.topLevelItem(top_index)
            manday = top_item.text(2)
            t_input = top_item.text(3)
            dm = top_item.text(4)
            lay = top_item.text(5)
            ani = top_item.text(6)
            cfx = top_item.text(7)
            fx = top_item.text(8)
            lig = top_item.text(9)
            mp = top_item.text(10)
            roto = top_item.text(11)
            paint = top_item.text(12)
            comp = top_item.text(13)
            output = top_item.text(14)
            if manday != '':
                total_manday.append(float(manday))
            if t_input != '':
                total_input.append(float(t_input))
            if dm != '':
                total_dm.append(float(dm))
            if lay != '':
                total_lay.append(float(lay))
            if ani != '':
                total_ani.append(float(ani))
            if cfx != '':
                total_cfx.append(float(cfx))
            if fx != '':
                total_fx.append(float(fx))
            if lig != '':
                total_lig.append(float(lig))
            if mp != '':
                total_mp.append(float(mp))
            if roto != '':
                total_roto.append(float(roto))
            if paint != '':
                total_paint.append(float(paint))
            if comp != '':
                total_comp.append(float(comp))
            if output != '':
                total_output.append(float(output))
        
        self.all_sum_item = QTreeWidgetItem(self.main_table)
        self.all_sum_item.setText(0,'Sum Man-day')
        self.all_sum_item.setText(2,"%.1f"%(sum(total_manday)))
        self.all_sum_item.setText(3,"%.1f"%(sum(total_input)))
        self.all_sum_item.setText(4,"%.1f"%(sum(total_dm)))
        self.all_sum_item.setText(5,"%.1f"%(sum(total_lay)))
        self.all_sum_item.setText(6,"%.1f"%(sum(total_ani)))
        self.all_sum_item.setText(7,"%.1f"%(sum(total_cfx)))
        self.all_sum_item.setText(8,"%.1f"%(sum(total_fx)))
        self.all_sum_item.setText(9,"%.1f"%(sum(total_lig)))
        self.all_sum_item.setText(10,"%.1f"%(sum(total_mp)))
        self.all_sum_item.setText(11,"%.1f"%(sum(total_roto)))
        self.all_sum_item.setText(12,"%.1f"%(sum(total_paint)))
        self.all_sum_item.setText(13,"%.1f"%(sum(total_comp)))
        self.all_sum_item.setText(14,"%.1f"%(sum(total_output)))
        self.all_sum_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
        self.all_sum_item.setTextAlignment(14,QtCore.Qt.AlignCenter)

    def genera_main_by_project(self):
        current_year = self.comb_year.currentText()
        
        self.main_table.clear()
        self.header = ['Task','Artist','Total','input','dm','layout','anim','cfx','fx','lighting','mp','roto','paint','comp','output']
        self.main_table.header().setStretchLastSection(False)
        self.main_table.setHeaderLabels(self.header)
        for header_index in range(len(self.header)):
            if header_index == 0:
                self.main_table.setColumnWidth(header_index,200)
            else:
                self.main_table.setColumnWidth(header_index,75)
                
        current_project = self.list_user.selectedItems()
        for i in current_project:
            select_proj = i.text(0)
            
        info = cgt_info.info.get_task_manday('proj_%s'%(select_proj.lower()))
        
        seq_list = []
        for artist in info:
            if artist != '待分配' and artist != '':
                for project in info[artist]:
                    for seq in info[artist][project]:
                        for task in info[artist][project][seq]:
                            pipeline_step = info[artist][project][seq][task]['pipline_step']
                            seq_list.append({
                                seq:{
                                    pipeline_step:{
                                        artist:'%s|%s'%(task,info[artist][project][seq][task]['mandays'])
                                    }
                                    
                                }
                            })
                        
        reslut_seq = {}
        reslut_pipeline = {}
        
        for index in seq_list:
            for seq_k,seq_v in index.items():
                reslut_seq.setdefault(seq_k,[]).append(seq_v)
                
        for seq_index in reslut_seq:
            pipe_dict = {}
            for art_index in reslut_seq[seq_index]:
                for pipe_k,pipe_v in art_index.items():
                    pipe_dict.setdefault(pipe_k,[]).append(pipe_v)
            reslut_pipeline[seq_index] = pipe_dict

        for seq_name in reslut_pipeline:
            seq_item = QTreeWidgetItem(self.main_table)
            seq_item.setText(0,seq_name)
            for pipe_name in reslut_pipeline[seq_name]:
                set_index = self.header.index(pipe_name)
                for tasks in reslut_pipeline[seq_name][pipe_name]:
                    for task_index in tasks:
                        task_item = QTreeWidgetItem(seq_item)
                        task_name = tasks[task_index].split('|')[0]
                        task_manday = tasks[task_index].split('|')[1]
                        task_item.setText(0,task_name)
                        task_item.setText(1,task_index)
                        task_item.setText(set_index,task_manday)
                        task_item.setTextAlignment(1,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
                        task_item.setTextAlignment(14,QtCore.Qt.AlignCenter)
        
        top_count = self.main_table.topLevelItemCount()
        for top_index in range(top_count):
            sum_input = []
            sum_dm = []
            sum_layout = []
            sum_anim = []
            sum_cfx = []
            sum_fx = []
            sum_lig = []
            sum_mp = []
            sum_roto = []
            sum_paint = []
            sum_comp = []
            sum_output = []
            total_manday = []
            top_item = self.main_table.topLevelItem(top_index)
            top_item.setTextAlignment(2,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(3,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(4,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(5,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(6,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(7,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(8,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(9,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(10,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(11,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(12,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(13,QtCore.Qt.AlignCenter)
            top_item.setTextAlignment(14,QtCore.Qt.AlignCenter)
            child_count = top_item.childCount()
            for child_index in range(child_count):
                child_item = top_item.child(child_index)
                input_manday = child_item.text(3)
                dm_manday = child_item.text(4)
                layout_manday = child_item.text(5)
                anim_manday = child_item.text(6)
                cfx_manday = child_item.text(7)
                fx_manday = child_item.text(8)
                lig_manday = child_item.text(9)
                mp_manday = child_item.text(10)
                roto_manday = child_item.text(11)
                panit_manday = child_item.text(12)
                comp_manday = child_item.text(13)
                output_manday = child_item.text(14)
                if input_manday != '':
                    sum_input.append(float(input_manday))
                else:
                    pass
                
                if dm_manday != '':
                    sum_dm.append(float(dm_manday))
                else:
                    pass
                
                if layout_manday != '':
                    sum_layout.append(float(layout_manday))
                else:
                    pass
                
                if anim_manday != '':
                    sum_anim.append(float(anim_manday))
                else:
                    pass
                
                if cfx_manday != '':
                    sum_cfx.append(float(cfx_manday))
                else:
                    pass
                
                if fx_manday != '':
                    sum_fx.append(float(fx_manday))
                else:
                    pass
                
                if lig_manday != '':
                    sum_lig.append(float(lig_manday))
                else:
                    pass
                
                if mp_manday != '':
                    sum_mp.append(float(mp_manday))
                else:
                    pass
                
                if roto_manday != '':
                    sum_roto.append(float(roto_manday))
                else:
                    pass
                
                if panit_manday != '':
                    sum_paint.append(float(panit_manday))
                else:
                    pass
                
                if comp_manday != '':
                    sum_comp.append(float(comp_manday))
                else:
                    pass
                
                if output_manday != '':
                    sum_output.append(float(output_manday))
                else:
                    pass
                
            if len(sum_input) > 0:
                sum_input_num = sum(sum_input)
                total_manday.append(sum_input_num)
                top_item.setText(3,"%.1f"%(sum_input_num))
            else:
                pass
            if len(sum_dm) > 0:
                sum_dm_num = sum(sum_dm)
                total_manday.append(sum_dm_num)
                top_item.setText(4,"%.1f"%(sum_dm_num))
            else:
                pass
            if len(sum_layout) > 0:
                sum_layout_num = sum(sum_layout)
                total_manday.append(sum_layout_num)
                top_item.setText(5,"%.1f"%(sum_layout_num))
            else:
                pass
            if len(sum_anim) > 0:
                sum_anim_num = sum(sum_anim)
                total_manday.append(sum_anim_num)
                top_item.setText(6,"%.1f"%(sum_anim_num))
            else:
                pass
            if len(sum_cfx) > 0:
                sum_cfx_num = sum(sum_cfx)
                total_manday.append(sum_cfx_num)
                top_item.setText(7,"%.1f"%(sum_cfx_num))
            else:
                pass
            if len(sum_fx) > 0:
                sum_fx_num = sum(sum_fx)
                total_manday.append(sum_fx_num)
                top_item.setText(8,"%.1f"%(sum_fx_num))
            else:
                pass
            if len(sum_lig) > 0:
                sum_lig_num = sum(sum_lig)
                total_manday.append(sum_lig_num)
                top_item.setText(9,"%.1f"%(sum_lig_num))
            else:
                pass
            if len(sum_mp) > 0:
                sum_mp_num = sum(sum_mp)
                total_manday.append(sum_mp_num)
                top_item.setText(10,"%.1f"%(sum_mp_num))
            else:
                pass
            if len(sum_roto) > 0:
                sum_roto_num = sum(sum_roto)
                total_manday.append(sum_roto_num)
                top_item.setText(11,"%.1f"%(sum_roto_num))
            else:
                pass
            if len(sum_paint) > 0:
                sum_paint_num = sum(sum_paint)
                total_manday.append(sum_paint_num)
                top_item.setText(12,"%.1f"%(sum_paint_num))
            else:
                pass
            if len(sum_comp) > 0:
                sum_comp_num = sum(sum_comp)
                total_manday.append(sum_comp_num)
                top_item.setText(13,"%.1f"%(sum_comp_num))
            else:
                pass
            if len(sum_output) > 0:
                sum_output_num = sum(sum_output)
                total_manday.append(sum_output_num)
                top_item.setText(14,"%.1f"%(sum_output_num))
            else:
                pass
            if len(total_manday) > 0:
                total = sum(total_manday)
                top_item.setText(2,"%.1f"%(total))
                
        self.genera_total_item_by_project()
                
        main_it = QTreeWidgetItemIterator(self.main_table)
        while main_it.value():
            for col in range(len(self.header)):
                item_frame = QFrame()
                item_frame.setFrameShape(QFrame.Box)
                item_frame.setStyleSheet("QFrame{border:0.5px solid #E5E5E5};")
                self.main_table.setItemWidget(main_it.value(),col,item_frame)
            main_it += 1
    
    def refresh(self):
        current_year = self.comb_year.currentText()
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        self.main_table.clear()

        if self.radbtn_artist.isChecked():
            self.list_user.clear()
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.users = json.load(f)
                
            json_path = '%s/temp_cgtdata.json'%(json_root)
            tasks = cgt_info.info.get_tasks_manday()
            with open(json_path,'w') as f:
                json.dump(tasks,f)
                        
            self.get_user_list()
            self.generate_main_by_artist()
        if self.radbtn_project.isChecked():
            current_project = self.list_user.selectedItems()
            for i in current_project:
                select_proj = i.text(0)
                
            self.genera_main_by_project()
            
    def fliter_project(self):
        check_hide = []
        current_project = self.combox_project.currentText()
        fliter_project = current_project.split(',')
        
        top_count = self.main_table.topLevelItemCount()
        for top_index in range(top_count):
            artist_item = self.main_table.topLevelItem(top_index)
            children_count = artist_item.childCount()
            for child_index in range(children_count):
                child_item = artist_item.child(child_index)
                child_count = child_item.childCount()
                if child_count > 0:
                    project = child_item.text(0)
                    if project not in fliter_project:
                        child_item.setHidden(1)
                    else:
                        child_item.setHidden(0)
                else:
                    pass
                
    def sum_show_total(self):
        self.all_sum_item.setText(1,'0.0')
        self.all_sum_item.setText(2,'0.0')
        self.all_sum_item.setText(3,'0.0')
        self.all_sum_item.setText(4,'0.0')
        self.all_sum_item.setText(5,'0.0')
        self.all_sum_item.setText(6,'0.0')
        self.all_sum_item.setText(7,'0.0')
        self.all_sum_item.setText(8,'0.0')
        self.all_sum_item.setText(9,'0.0')
        self.all_sum_item.setText(10,'0.0')
        self.all_sum_item.setText(11,'0.0')
        self.all_sum_item.setText(12,'0.0')
        self.all_sum_item.setText(13,'0.0')
        
        total_manday = []
        jan_manday = []
        feb_manday = []
        mar_manday = []
        apr_manday = []
        may_manday = []
        jun_manday = []
        jul_manday = []
        aug_manday = []
        sept_manday = []
        oct_manday = []
        nov_manday = []
        dec_manday = []
        for top_show_index in range(self.main_table.topLevelItemCount()):
            top_item = self.main_table.topLevelItem(top_show_index)
            if top_item.isHidden() == 0 and top_item.isDisabled() == 0 and top_item.text(0) != 'Sum Man-day':
                total_manday_value = top_item.text(1)
                total_jan_value = top_item.text(2)
                total_feb_value = top_item.text(3)
                total_mar_value = top_item.text(4)
                total_apr_value = top_item.text(5)
                total_may_value = top_item.text(6)
                total_jun_value = top_item.text(7)
                total_jul_value = top_item.text(8)
                total_aug_value = top_item.text(9)
                total_sept_value = top_item.text(10)
                total_oct_value = top_item.text(11)
                total_nov_value = top_item.text(12)
                total_dec_value = top_item.text(13)
                if total_manday_value != '':
                    total_manday.append(float(total_manday_value))
                else:
                    total_manday.append(0.0)
                if total_jan_value != '':
                    jan_manday.append(float(total_jan_value))
                else:
                    jan_manday.append(0.0)
                if total_feb_value != '':
                    feb_manday.append(float(total_feb_value))
                else:
                    feb_manday.append(0.0)
                if total_mar_value != '':
                    mar_manday.append(float(total_mar_value))
                else:
                    mar_manday.append(0.0)
                if total_apr_value != '':
                    apr_manday.append(float(total_apr_value))
                else:
                    apr_manday.append(0.0)
                if total_may_value != '':
                    may_manday.append(float(total_may_value))
                else:
                    may_manday.append(0.0)
                if total_jun_value != '':
                    jun_manday.append(float(total_jun_value))
                else:
                    jun_manday.append(0.0)
                if total_jul_value != '':
                    jul_manday.append(float(total_jul_value))
                else:
                    jul_manday.append(0.0)
                if total_aug_value != '':
                    aug_manday.append(float(total_aug_value))
                else:
                    aug_manday.append(0.0)
                if total_sept_value != '':
                    sept_manday.append(float(total_sept_value))
                else:
                    sept_manday.append(0.0)
                if total_oct_value != '':
                    oct_manday.append(float(total_oct_value))
                else:
                    oct_manday.append(0.0)
                if total_nov_value != '':
                    nov_manday.append(float(total_nov_value))
                else:
                    nov_manday.append(0.0)
                if total_dec_value != '':
                    dec_manday.append(float(total_dec_value))
                else:
                    dec_manday.append(0.0)
                
                self.all_sum_item.setText(1,str("%.1f"%(sum(total_manday))))
                self.all_sum_item.setText(2,str("%.1f"%(sum(jan_manday))))
                self.all_sum_item.setText(3,str("%.1f"%(sum(feb_manday))))
                self.all_sum_item.setText(4,str("%.1f"%(sum(mar_manday))))
                self.all_sum_item.setText(5,str("%.1f"%(sum(apr_manday))))
                self.all_sum_item.setText(6,str("%.1f"%(sum(may_manday))))
                self.all_sum_item.setText(7,str("%.1f"%(sum(jun_manday))))
                self.all_sum_item.setText(8,str("%.1f"%(sum(jul_manday))))
                self.all_sum_item.setText(9,str("%.1f"%(sum(aug_manday))))
                self.all_sum_item.setText(10,str("%.1f"%(sum(sept_manday))))
                self.all_sum_item.setText(11,str("%.1f"%(sum(oct_manday))))
                self.all_sum_item.setText(12,str("%.1f"%(sum(nov_manday))))
                self.all_sum_item.setText(13,str("%.1f"%(sum(dec_manday))))
            else:
                pass
                
    def showContextMenu_maintable(self,pos):
        self.contexMenu_maintable = QtWidgets.QMenu()
        self.contexMenu_maintable.clear()
        expand_menu = self.contexMenu_maintable.addMenu('Expand')
        self.contexMenu_maintable.addSeparator()
        collapse_menu = self.contexMenu_maintable.addMenu('collapse')
        self.contexMenu_maintable.addSeparator()
        
        expand_all = expand_menu.addAction('Expand All')
        expand_menu.addSeparator()
        expandCurrentSelect = expand_menu.addAction('Expand Current Select')
        expand_menu.addSeparator()
        
        collapse_all = collapse_menu.addAction('Collapse All')
        collapse_menu.addSeparator()
        collapseCurrentSelect = collapse_menu.addAction('Collapse Current Select')
        collapse_menu.addSeparator()
        
        hide_select = self.contexMenu_maintable.addAction('Hide Current Select')
        self.contexMenu_maintable.addSeparator()
        show_all_item = self.contexMenu_maintable.addAction('Show All Item')
        self.contexMenu_maintable.addSeparator()
        
        expand_all.triggered.connect(self.maintable_expandAll)
        collapse_all.triggered.connect(self.maintable_collapse)
        expandCurrentSelect.triggered.connect(self.maintable_expandCurrentSelect)
        collapseCurrentSelect.triggered.connect(self.maintable_collapseCurrentSelect)
        hide_select.triggered.connect(self.maintable_hideSelect)
        show_all_item.triggered.connect(self.maintable_showAllItem)
        
        if self.main_table.itemAt(pos):
            self.contexMenu_maintable.move(QtGui.QCursor().pos())
            self.contexMenu_maintable.show()
            
    def maintable_expandAll(self):
        self.main_table.expandAll()
        
    def maintable_collapse(self):
        self.main_table.collapseAll()
        
    def maintable_expandCurrentSelect(self):
        current_item = self.main_table.currentItem()
        current_item.setExpanded(1)
        
    def maintable_collapseCurrentSelect(self):
        current_item = self.main_table.currentItem()
        current_item.setExpanded(0)
        
    def maintable_expandCurrentSelect_project(self):
        current_item = self.main_table.currentItem()
        
        if current_item.parent() == None:
            current_item.setExpanded(1)
        else:
            pass
        
    def maintable_collapseCurrentSelcet_project(self):
        current_item = self.main_table.currentItem()
        
        if current_item != None:
            print(current_item.text(0))
        else:
            pass
            
    def maintable_hideSelect(self):
        current_item = self.main_table.currentItem()
        
        if current_item.parent() != None:
            if current_item.parent().parent() != None:
                current_top_text = current_item.parent().parent().text(0)
                user_group_count = self.list_user.topLevelItemCount()
                for group_index in range(user_group_count):
                    group_item = self.list_user.topLevelItem(group_index)
                    for child_index in range(group_item.childCount()):
                        child_item = group_item.child(child_index)
                        if child_item.text(0) == current_top_text:
                            if child_item.checkState(0) == Qt.Checked:
                                child_item.setCheckState(0,Qt.Unchecked)
                
                current_item.parent().parent().setHidden(1)
            else:
                current_top_text = current_item.parent().text(0)
                user_group_count = self.list_user.topLevelItemCount()
                for group_index in range(user_group_count):
                    group_item = self.list_user.topLevelItem(group_index)
                    for child_index in range(group_item.childCount()):
                        child_item = group_item.child(child_index)
                        if child_item.text(0) == current_top_text:
                            if child_item.checkState(0) == Qt.Checked:
                                child_item.setCheckState(0,Qt.Unchecked)
                
                current_item.parent().setHidden(1)
        else:
            current_text = current_item.text(0)
            user_group_count = self.list_user.topLevelItemCount()
            for group_index in range(user_group_count):
                group_item = self.list_user.topLevelItem(group_index)
                for child_index in range(group_item.childCount()):
                    child_item = group_item.child(child_index)
                    if child_item.text(0) == current_text:
                        if child_item.checkState(0) == Qt.Checked:
                            child_item.setCheckState(0,Qt.Unchecked)
            
            current_item.setHidden(1)
        
        self.sum_show_total()
            
    def maintable_showAllItem(self):
        top = self.main_table.topLevelItemCount()
        for index in range(top):
            item = self.main_table.topLevelItem(index)
            item.setHidden(0)
            
        user_group_count = self.list_user.topLevelItemCount()
        for group_index in range(user_group_count):
            group_item = self.list_user.topLevelItem(group_index)
            for child_index in range(group_item.childCount()):
                child_item = group_item.child(child_index)
                if child_item.isDisabled() == False:
                    if child_item.checkState(0) == Qt.Unchecked:
                        child_item.setCheckState(0,Qt.Checked)
                        
        self.sum_show_total()
        
    def fliter_artist(self):
        user_count = self.list_user.count()
        check_list = []
        for row_index in range(user_count):
            item = self.list_user.item(row_index)
            if item.checkState() == Qt.Checked:
                check_list.append(item.text())
                
        top_count = self.main_table.topLevelItemCount()
        for top_index in range(top_count):
            artist_item = self.main_table.topLevelItem(top_index)
            if artist_item.text(0) not in check_list:
                artist_item.setHidden(1)
            else:
                pass
            
    def showContextMenu_user(self,pos):
        if self.radbtn_artist.isChecked():
            self.contexMenu_user = QtWidgets.QMenu()
            self.contexMenu_user.clear()
            select_all = self.contexMenu_user.addAction('Select All')
            self.contexMenu_user.addSeparator()
            deselect_all = self.contexMenu_user.addAction('Deselect All')
            self.contexMenu_user.addSeparator()
            select_group = self.contexMenu_user.addMenu('Select Group')
            select_prod = select_group.addAction('Select Prod')
            select_group.addSeparator()
            select_fx = select_group.addAction('Select FX')
            select_group.addSeparator()
            select_comp = select_group.addAction('Select Comp')
            select_group.addSeparator()
            select_anim = select_group.addAction('Select Anim')
            select_group.addSeparator()
            select_td = select_group.addAction('Select TD')
            self.contexMenu_user.addSeparator()
            deselect_group = self.contexMenu_user.addMenu('Deselect Group')
            deselect_prod = deselect_group.addAction('Deselect Prod')
            deselect_group.addSeparator()
            deselect_fx = deselect_group.addAction('Deselect FX')
            deselect_group.addSeparator()
            deselect_comp = deselect_group.addAction('Deselect Comp')
            deselect_group.addSeparator()
            deselect_anim = deselect_group.addAction('Deselect Anim')
            deselect_group.addSeparator()
            deselect_td = deselect_group.addAction('Deselect TD')
            
            self.contexMenu_user.addSeparator()
            
            select_all.triggered.connect(self.user_select_all)
            deselect_all.triggered.connect(self.user_deselect_all)
            select_prod.triggered.connect(self.user_select_prod)
            select_fx.triggered.connect(self.user_select_fx)
            select_comp.triggered.connect(self.user_select_comp)
            select_anim.triggered.connect(self.user_select_anim)
            select_td.triggered.connect(self.user_select_td)
            deselect_prod.triggered.connect(self.user_deselect_prod)
            deselect_fx.triggered.connect(self.user_deselect_fx)
            deselect_comp.triggered.connect(self.user_deselect_comp)
            deselect_anim.triggered.connect(self.user_deselect_anim)
            deselect_td.triggered.connect(self.user_deselect_td)
            
            if self.list_user.itemAt(pos):
                self.contexMenu_user.move(QtGui.QCursor().pos())
                self.contexMenu_user.show()
        else:
            pass
        
    def user_select_all(self):
        self.list_user.expandAll()
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.checkState(0) == Qt.Unchecked:
                top_item.setCheckState(0,Qt.Checked)
            top_item.setExpanded(1)
            child_count = self.list_user.topLevelItem(top_index).childCount()
            for child_index in range(child_count):
                child_item = self.list_user.topLevelItem(top_index).child(child_index)
                if child_item.isDisabled() == False:
                    if child_item.checkState(0) == Qt.Unchecked:
                        child_item.setCheckState(0,Qt.Checked)
                        check_artist.append(child_item.text(0))
                    else:
                        check_artist.append(child_item.text(0))
                else:
                    pass
        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            artist_item.setHidden(0)
            
        self.sum_show_total()
        
    def user_deselect_all(self):
        self.list_user.collapseAll()
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.checkState(0) == Qt.Checked:
                top_item.setCheckState(0,Qt.Unchecked)
            top_item.setExpanded(0)
            child_count = self.list_user.topLevelItem(top_index).childCount()
            for child_index in range(child_count):
                child_item = self.list_user.topLevelItem(top_index).child(child_index)
                if child_item.isDisabled() == False:
                    if child_item.checkState(0) == Qt.Checked:
                        child_item.setCheckState(0,Qt.Unchecked)
                        check_artist.append(child_item.text(0))
                    else:
                        check_artist.append(child_item.text(0))
                else:
                    pass
        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) != 'Sum Man-day':
                artist_item.setHidden(1)
            else:
                pass
            
        self.sum_show_total()
    
    def user_select_prod(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'Prod':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                prod_count = self.list_user.topLevelItem(top_index).childCount()
                for prod_index in range(prod_count):
                    prod_item = self.list_user.topLevelItem(top_index).child(prod_index)
                    if prod_item.isDisabled() == False:
                        if prod_item.checkState(0) == Qt.Unchecked:
                            prod_item.setCheckState(0,Qt.Checked)
                            check_artist.append(prod_item.text(0))
                        else:
                            check_artist.append(prod_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(0)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_select_fx(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'FX':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                fx_count = self.list_user.topLevelItem(top_index).childCount()
                for fx_index in range(fx_count):
                    fx_item = self.list_user.topLevelItem(top_index).child(fx_index)
                    if fx_item.isDisabled() == False:
                        if fx_item.checkState(0) == Qt.Unchecked:
                            fx_item.setCheckState(0,Qt.Checked)
                            check_artist.append(fx_item.text(0))
                        else:
                            check_artist.append(fx_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) != 'Sum Man-day':
                if artist_item.text(0) in check_artist:
                    artist_item.setHidden(0)
                else:
                    pass
            else:
                artist_item.setHidden(0)
            
        self.sum_show_total()
                        
    def user_select_comp(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'Comp':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                comp_count = self.list_user.topLevelItem(top_index).childCount()
                for comp_index in range(comp_count):
                    comp_item = self.list_user.topLevelItem(top_index).child(comp_index)
                    if comp_item.isDisabled() == False:
                        if comp_item.checkState(0) == Qt.Unchecked:
                            comp_item.setCheckState(0,Qt.Checked)
                            check_artist.append(comp_item.text(0))
                        else:
                            check_artist.append(comp_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(0)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_select_anim(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'Anim':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                anim_count = self.list_user.topLevelItem(top_index).childCount()
                for anim_index in range(anim_count):
                    anim_item = self.list_user.topLevelItem(top_index).child(anim_index)
                    if anim_item.isDisabled() == False:
                        if anim_item.checkState(0) == Qt.Unchecked:
                            anim_item.setCheckState(0,Qt.Checked)
                            check_artist.append(anim_item.text(0))
                        else:
                            check_artist.append(anim_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(0)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_select_td(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'TD':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                td_count = self.list_user.topLevelItem(top_index).childCount()
                for td_index in range(td_count):
                    td_item = self.list_user.topLevelItem(top_index).child(td_index)
                    if td_item.isDisabled() == False:
                        if td_item.checkState(0) == Qt.Unchecked:
                            td_item.setCheckState(0,Qt.Checked)
                            check_artist.append(td_item.text(0))
                        else:
                            check_artist.append(td_item.text(0))
                    else:
                        pass

        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(0)
            else:
                pass
            
        self.sum_show_total()
            
    def user_deselect_prod(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'Prod':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                prod_count = self.list_user.topLevelItem(top_index).childCount()
                for prod_index in range(prod_count):
                    prod_item = self.list_user.topLevelItem(top_index).child(prod_index)
                    if prod_item.isDisabled() == False:
                        if prod_item.checkState(0) == Qt.Checked:
                            prod_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(prod_item.text(0))
                        else:
                            check_artist.append(prod_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(1)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_deselect_fx(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'FX':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                fx_count = self.list_user.topLevelItem(top_index).childCount()
                for fx_index in range(fx_count):
                    fx_item = self.list_user.topLevelItem(top_index).child(fx_index)
                    if fx_item.isDisabled() == False:
                        if fx_item.checkState(0) == Qt.Checked:
                            fx_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(fx_item.text(0))
                        else:
                            check_artist.append(fx_item.text(0))
                    else:
                        pass
                
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(1)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_deselect_comp(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'Comp':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                comp_count = self.list_user.topLevelItem(top_index).childCount()
                for comp_index in range(comp_count):
                    comp_item = self.list_user.topLevelItem(top_index).child(comp_index)
                    if comp_item.isDisabled() == False:
                        if comp_item.checkState(0) == Qt.Checked:
                            comp_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(comp_item.text(0))
                        else:
                            check_artist.append(comp_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(1)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_deselect_anim(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'Anim':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                anim_count = self.list_user.topLevelItem(top_index).childCount()
                for anim_index in range(anim_count):
                    anim_item = self.list_user.topLevelItem(top_index).child(anim_index)
                    if anim_item.isDisabled() == False:
                        if anim_item.checkState(0) == Qt.Checked:
                            anim_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(anim_item.text(0))
                        else:
                            check_artist.append(anim_item.text(0))
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(1)
            else:
                pass
            
        self.sum_show_total()
                        
    def user_deselect_td(self):
        top_count = self.list_user.topLevelItemCount()
        check_artist = []
        for top_index in range(top_count):
            top_item = self.list_user.topLevelItem(top_index)
            if top_item.text(0) == 'TD':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                td_count = self.list_user.topLevelItem(top_index).childCount()
                for td_index in range(td_count):
                    td_item = self.list_user.topLevelItem(top_index).child(td_index)
                    if td_item.isDisabled() == False:
                        if td_item.checkState(0) == Qt.Checked:
                            td_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(td_item.text(0))
                        else:
                            check_artist.append(td_item.text(0))
                    else:
                        pass
                        
        artist_count = self.main_table.topLevelItemCount()
        for artist_index in range(artist_count):
            artist_item = self.main_table.topLevelItem(artist_index)
            if artist_item.text(0) in check_artist:
                artist_item.setHidden(1)
            else:
                pass
            
        self.sum_show_total()
            
    def check_user_stage(self):
        if self.radbtn_artist.isChecked():
            check_artist = []
            top_count = self.list_user.topLevelItemCount()
            artist_count = self.main_table.topLevelItemCount()
            if self.checkbox_only.isChecked():
                self.list_user.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
                current_select = self.list_user.selectedItems()
                for select_index in current_select:
                    artist_text = select_index.text(0)
                    check_artist.append(artist_text)
                    
                for artist_index in range(artist_count):
                    artist_item = self.main_table.topLevelItem(artist_index)
                    if artist_item.text(0) != 'Sum Man-day':
                        if artist_item.text(0) in check_artist:
                            artist_item.setHidden(0)
                        else:
                            artist_item.setHidden(1)
                    else:
                        artist_item.setHidden(0)
                    
            else:
                for artist_index in range(artist_count):
                    artist_item = self.main_table.topLevelItem(artist_index)
                    artist_item.setHidden(0)
            
                for top_index in range(top_count):
                    top_item = self.list_user.topLevelItem(top_index)
                    child_count = top_item.childCount()
                    for child_index in range(child_count):
                        check_item = top_item.child(child_index)
                        if check_item.isDisabled() == False:
                            if check_item.checkState(0) == Qt.Checked:
                                check_artist.append(check_item.text(0))
                            
                for artist_index in range(artist_count):
                    artist_item = self.main_table.topLevelItem(artist_index)
                    if artist_item.text(0) != 'Sum Man-day':
                        if artist_item.text(0) not in check_artist:
                            artist_item.setHidden(1)
                        else:
                            artist_item.setHidden(0)
                    else:
                        artist_item.setHidden(0)
                        
                check_artist.clear()
                self.list_user.clearSelection()
                
            self.sum_show_total()
        else:
            pass
        
    def checkbox_stage(self):
        top_count = self.list_user.topLevelItemCount()
        artist_count = self.main_table.topLevelItemCount()
        check_artist = []
        
        if self.checkbox_only.isChecked():
            for artist_index in range(artist_count):
                artist_item = self.main_table.topLevelItem(artist_index)
                if artist_item.isDisabled() == 0:
                    artist_item.setHidden(1)
                else:
                    artist_item.setHidden(0)
        else:       
            self.list_user.clearSelection()
            for artist_index in range(artist_count):
                artist_item = self.main_table.topLevelItem(artist_index)
                artist_item.setHidden(0)
                
            for top_index in range(top_count):
                top_item = self.list_user.topLevelItem(top_index)
                child_count = top_item.childCount()
                for child_index in range(child_count):
                    check_item = top_item.child(child_index)
                    if check_item.isDisabled() == False:
                        if check_item.checkState(0) == Qt.Checked:
                            check_artist.append(check_item.text(0))
                    # else:
                    #     check_item.setCheckState(0,Qt.Unchecked)
                        
            for check_index in range(artist_count):
                check_item = self.main_table.topLevelItem(check_index)
                if check_item.text(0) != 'Sum Man-day':
                    if check_item.text(0) not in check_artist:
                        check_item.setHidden(1)
                    else:
                        check_item.setHidden(0)
                else:
                    check_item.setHidden(0)
                    
            check_artist.clear()
            
        self.sum_show_total()
                    
    def show_export(self):
        artist_info = self.main_table.topLevelItemCount()
        info = []
        for index in range(artist_info):
            artist = self.main_table.topLevelItem(index).text(0)
            project_count = self.main_table.topLevelItem(index).childCount()
            for project_index in range(project_count):
                if self.main_table.topLevelItem(index).child(project_index).childCount() != 0:
                    project_item = self.main_table.topLevelItem(index).child(project_index)
                    seq_count = project_item.childCount()
                    for seq_index in range(seq_count):
                        seq_item = project_item.child(seq_index)
                        sequence = seq_item.text(0)
                        shot_count = seq_item.childCount()
                        for shot_index in range(shot_count):
                            shot_item = seq_item.child(shot_index)
                            shot_name = shot_item.text(0)
                            shot_manday = shot_item.text(1)
                            shot_jan = shot_item.text(2)
                            shot_feb = shot_item.text(3)
                            shot_mar = shot_item.text(4)
                            shot_apr = shot_item.text(5)
                            shot_may = shot_item.text(6)
                            shot_jun = shot_item.text(7)
                            shot_jul = shot_item.text(8)
                            shot_aug = shot_item.text(9)
                            shot_sept = shot_item.text(10)
                            shot_oct = shot_item.text(11)
                            shot_nov = shot_item.text(12)
                            shot_dec = shot_item.text(13)
                            shot_info = [shot_name,shot_manday,shot_jan,shot_feb,shot_mar,shot_apr,shot_may,shot_jun,shot_jul,shot_aug,shot_sept,shot_oct,shot_nov,shot_dec]

                            project = self.main_table.topLevelItem(index).child(project_index).text(0)
                            manday = seq_item.text(1)
                            jan = seq_item.text(2)
                            feb = seq_item.text(3)
                            mar = seq_item.text(4)
                            apr = seq_item.text(5)
                            may = seq_item.text(6)
                            jun = seq_item.text(7)
                            jul = seq_item.text(8)
                            aug = seq_item.text(9)
                            sept = seq_item.text(10)
                            oct = seq_item.text(11)
                            nov = seq_item.text(12)
                            dec = seq_item.text(13)
                            info.append('%s-%s-%s-%s-%s-%s-%s-%s-%s-%s-%s-%s-%s-%s-%s-%s|%s'%(artist,project,sequence,manday,jan,feb,mar,apr,may,jun,jul,aug,sept,oct,nov,dec,shot_info))
                
        self.export_show = export_excel(info)
        self.export_show.show()
                
class export_excel(QMainWindow):
    
    export_excel_Signal = QtCore.Signal(list)
    
    def __init__(self,info):
        super(export_excel,self).__init__()
        self.data_info = info
        self.resize(600,600)
        self.setWindowTitle(u'Export Excel')
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        
        self.lab_artist = QLabel('Artist')
        self.list_artist = QTreeWidget()
        self.list_artist.setHeaderHidden(1)
        self.lab_project = QLabel('Project')
        self.list_project = QListWidget()
        self.checkbox_artist = QCheckBox('Artist')
        self.checkbox_artist.setChecked(1)
        self.checkbox_total = QCheckBox('Total')
        self.checkbox_total.setChecked(1)
        self.checkbox_shot = QCheckBox('Shot')
        self.checkbox_jan = QCheckBox('jan')
        self.checkbox_feb = QCheckBox('Feb')
        self.checkbox_mar = QCheckBox('Mar')
        self.checkbox_apr = QCheckBox('Apr')
        self.checkbox_may = QCheckBox('May')
        self.checkbox_jun = QCheckBox('Jun')
        self.checkbox_jul = QCheckBox('Jul')
        self.checkbox_aug = QCheckBox('Aug')
        self.checkbox_sept = QCheckBox('Sept')
        self.checkbox_oct = QCheckBox('Oct')
        self.checkbox_nov = QCheckBox('Nov')
        self.checkbox_dec = QCheckBox('Dec')
        self.lab_export_path = QLabel('Export Path:')
        self.line_export_path = QLineEdit()
        self.btn_serach = QPushButton('···')
        self.btn_export = QPushButton('Export')
        
        self.common_lay = QHBoxLayout()
        self.common_lay.addWidget(self.checkbox_artist)
        self.common_lay.addWidget(self.checkbox_total)
        self.common_lay.addWidget(self.checkbox_shot)
        
        self.artist_lay = QVBoxLayout()
        self.artist_lay.addWidget(self.lab_artist)
        self.artist_lay.addWidget(self.list_artist)
        
        self.project_lay = QVBoxLayout()
        self.project_lay.addWidget(self.lab_project)
        self.project_lay.addWidget(self.list_project)
        
        self.chose_lay = QHBoxLayout()
        self.chose_lay.addLayout(self.artist_lay)
        self.chose_lay.addLayout(self.project_lay)
        
        self.fir_season = QHBoxLayout()
        self.fir_season.addWidget(self.checkbox_jan)
        self.fir_season.addWidget(self.checkbox_feb)
        self.fir_season.addWidget(self.checkbox_mar)
        
        self.sec_season = QHBoxLayout()
        self.sec_season.addWidget(self.checkbox_apr)
        self.sec_season.addWidget(self.checkbox_may)
        self.sec_season.addWidget(self.checkbox_jun)
        
        self.thir_season = QHBoxLayout()
        self.thir_season.addWidget(self.checkbox_jul)
        self.thir_season.addWidget(self.checkbox_aug)
        self.thir_season.addWidget(self.checkbox_sept)
        
        self.forth_season = QHBoxLayout()
        self.forth_season.addWidget(self.checkbox_oct)
        self.forth_season.addWidget(self.checkbox_nov)
        self.forth_season.addWidget(self.checkbox_dec)
        
        self.month_lay = QVBoxLayout()
        self.month_lay.addLayout(self.common_lay)
        self.month_lay.addLayout(self.fir_season)
        self.month_lay.addLayout(self.sec_season)
        self.month_lay.addLayout(self.thir_season)
        self.month_lay.addLayout(self.forth_season)
        
        self.path_lay = QHBoxLayout()
        self.path_lay.addWidget(self.lab_export_path)
        self.path_lay.addWidget(self.line_export_path)
        self.path_lay.addWidget(self.btn_serach)

        self.showlay = QGridLayout()
        self.showlay.addLayout(self.chose_lay,0,0,1,3)
        self.showlay.addLayout(self.month_lay,0,4,1,1)
        self.showlay.addLayout(self.path_lay,2,0,1,1)
        self.showlay.addWidget(self.btn_export,2,4,1,1)
        
        self.childWindow = QWidget()
        self.childWindow.setLayout(self.showlay)
        self.setCentralWidget(self.childWindow)
        
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)

        try:
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.artist_info = json.load(f)
        except:
            print('user.json is not exist!')
            
        for group_index in self.artist_info:
            group_item = QTreeWidgetItem(self.list_artist)
            group_item.setText(0,group_index)
            group_item.setCheckState(0,Qt.Checked)
            group_item.setExpanded(1)
            group_item.setFlags(Qt.ItemFlag.ItemIsAutoTristate | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            for user_index in self.artist_info[group_index]:
                user_item = QTreeWidgetItem(group_item)
                user_item.setText(0,self.artist_info[group_index][user_index]['cn_name'])
                if self.artist_info[group_index][user_index]['check'] == 1:
                    user_item.setCheckState(0,Qt.Checked)
                else:
                    user_item.setDisabled(1)
                    user_item.setCheckState(0,Qt.Unchecked)
        
        projects = cgt_info.info.get_active_projects()
        projects.sort()
        for project_index in projects:
            project_item = QListWidgetItem(self.list_project)
            project_item.setCheckState(Qt.Checked)
            project_item.setText(project_index)
        
        self.list_artist.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.list_artist.customContextMenuRequested.connect(self.showContextMenu_user)
        self.btn_serach.clicked.connect(self.search_path)
        self.btn_export.clicked.connect(self.excel_export)
        
    def showContextMenu_user(self,pos):
        self.contexMenu_user = QtWidgets.QMenu()
        self.contexMenu_user.clear()
        select_all = self.contexMenu_user.addAction('Select All')
        self.contexMenu_user.addSeparator()
        deselect_all = self.contexMenu_user.addAction('Deselect All')
        self.contexMenu_user.addSeparator()
        select_group = self.contexMenu_user.addMenu('Select Group')
        select_prod = select_group.addAction('Select Prod')
        select_group.addSeparator()
        select_fx = select_group.addAction('Select FX')
        select_group.addSeparator()
        select_comp = select_group.addAction('Select Comp')
        select_group.addSeparator()
        select_anim = select_group.addAction('Select Anim')
        select_group.addSeparator()
        select_td = select_group.addAction('Select TD')
        self.contexMenu_user.addSeparator()
        deselect_group = self.contexMenu_user.addMenu('Deselect Group')
        deselect_prod = deselect_group.addAction('Deselect Prod')
        deselect_group.addSeparator()
        deselect_fx = deselect_group.addAction('Deselect FX')
        deselect_group.addSeparator()
        deselect_comp = deselect_group.addAction('Deselect Comp')
        deselect_group.addSeparator()
        deselect_anim = deselect_group.addAction('Deselect Anim')
        deselect_group.addSeparator()
        deselect_td = deselect_group.addAction('Deselect TD')
        
        self.contexMenu_user.addSeparator()
        
        select_all.triggered.connect(self.user_select_all)
        deselect_all.triggered.connect(self.user_deselect_all)
        select_prod.triggered.connect(self.user_select_prod)
        select_fx.triggered.connect(self.user_select_fx)
        select_comp.triggered.connect(self.user_select_comp)
        select_anim.triggered.connect(self.user_select_anim)
        select_td.triggered.connect(self.user_select_td)
        deselect_prod.triggered.connect(self.user_deselect_prod)
        deselect_fx.triggered.connect(self.user_deselect_fx)
        deselect_comp.triggered.connect(self.user_deselect_comp)
        deselect_anim.triggered.connect(self.user_deselect_anim)
        deselect_td.triggered.connect(self.user_deselect_td)
        
        if self.list_artist.itemAt(pos):
            self.contexMenu_user.move(QtGui.QCursor().pos())
            self.contexMenu_user.show()
        
    def user_select_all(self):
        self.list_artist.expandAll()
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.checkState(0) == Qt.Unchecked:
                top_item.setCheckState(0,Qt.Checked)
            top_item.setExpanded(1)
            child_count = self.list_artist.topLevelItem(top_index).childCount()
            for child_index in range(child_count):
                child_item = self.list_artist.topLevelItem(top_index).child(child_index)
                if child_item.isDisabled() == False:
                    if child_item.checkState(0) == Qt.Unchecked:
                        child_item.setCheckState(0,Qt.Checked)
                        check_artist.append(child_item.text(0))
                else:
                    pass
        
    def user_deselect_all(self):
        self.list_artist.collapseAll()
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.checkState(0) == Qt.Checked:
                top_item.setCheckState(0,Qt.Unchecked)
            top_item.setExpanded(0)
            child_count = self.list_artist.topLevelItem(top_index).childCount()
            for child_index in range(child_count):
                child_item = self.list_artist.topLevelItem(top_index).child(child_index)
                if child_item.isDisabled() == False:
                    if child_item.checkState(0) == Qt.Checked:
                        child_item.setCheckState(0,Qt.Unchecked)
                        check_artist.append(child_item.text(0))
                else:
                    pass
    
    def user_select_prod(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'Prod':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                prod_count = self.list_artist.topLevelItem(top_index).childCount()
                for prod_index in range(prod_count):
                    prod_item = self.list_artist.topLevelItem(top_index).child(prod_index)
                    if prod_item.isDisabled() == False:
                        if prod_item.checkState(0) == Qt.Unchecked:
                            prod_item.setCheckState(0,Qt.Checked)
                            check_artist.append(prod_item.text(0))
                    else:
                        pass
                        
    def user_select_fx(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'FX':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                fx_count = self.list_artist.topLevelItem(top_index).childCount()
                for fx_index in range(fx_count):
                    fx_item = self.list_artist.topLevelItem(top_index).child(fx_index)
                    if fx_item.isDisabled() == False:
                        if fx_item.checkState(0) == Qt.Unchecked:
                            fx_item.setCheckState(0,Qt.Checked)
                            check_artist.append(fx_item.text(0))
                    else:
                        pass
                        
    def user_select_comp(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'Comp':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                comp_count = self.list_artist.topLevelItem(top_index).childCount()
                for comp_index in range(comp_count):
                    comp_item = self.list_artist.topLevelItem(top_index).child(comp_index)
                    if comp_item.isDisabled() == False:
                        if comp_item.checkState(0) == Qt.Unchecked:
                            comp_item.setCheckState(0,Qt.Checked)
                            check_artist.append(comp_item.text(0))
                    else:
                        pass
                        
    def user_select_anim(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'Anim':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                anim_count = self.list_artist.topLevelItem(top_index).childCount()
                for anim_index in range(anim_count):
                    anim_item = self.list_artist.topLevelItem(top_index).child(anim_index)
                    if anim_item.isDisabled() == False:
                        if anim_item.checkState(0) == Qt.Unchecked:
                            anim_item.setCheckState(0,Qt.Checked)
                            check_artist.append(anim_item.text(0))
                    else:
                        pass
                        
    def user_select_td(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'TD':
                if top_item.checkState(0) == Qt.Unchecked:
                    top_item.setCheckState(0,Qt.Checked)
                else:
                    pass
                top_item.setExpanded(1)
                td_count = self.list_artist.topLevelItem(top_index).childCount()
                for td_index in range(td_count):
                    td_item = self.list_artist.topLevelItem(top_index).child(td_index)
                    if td_item.isDisabled() == False:
                        if td_item.checkState(0) == Qt.Unchecked:
                            td_item.setCheckState(0,Qt.Checked)
                            check_artist.append(td_item.text(0))
                    else:
                        pass
            
    def user_deselect_prod(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'Prod':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                prod_count = self.list_artist.topLevelItem(top_index).childCount()
                for prod_index in range(prod_count):
                    prod_item = self.list_artist.topLevelItem(top_index).child(prod_index)
                    if prod_item.isDisabled() == False:
                        if prod_item.checkState(0) == Qt.Checked:
                            prod_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(prod_item.text(0))
                    else:
                        pass
                        
    def user_deselect_fx(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'FX':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                fx_count = self.list_artist.topLevelItem(top_index).childCount()
                for fx_index in range(fx_count):
                    fx_item = self.list_artist.topLevelItem(top_index).child(fx_index)
                    if fx_item.isDisabled() == False:
                        if fx_item.checkState(0) == Qt.Checked:
                            fx_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(fx_item.text(0))
                    else:
                        pass
                        
    def user_deselect_comp(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'Comp':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                comp_count = self.list_artist.topLevelItem(top_index).childCount()
                for comp_index in range(comp_count):
                    comp_item = self.list_artist.topLevelItem(top_index).child(comp_index)
                    if comp_item.isDisabled() == False:
                        if comp_item.checkState(0) == Qt.Checked:
                            comp_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(comp_item.text(0))
                    else:
                        pass
                         
    def user_deselect_anim(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'Anim':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                anim_count = self.list_artist.topLevelItem(top_index).childCount()
                for anim_index in range(anim_count):
                    anim_item = self.list_artist.topLevelItem(top_index).child(anim_index)
                    if anim_item.isDisabled() == False:
                        if anim_item.checkState(0) == Qt.Checked:
                            anim_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(anim_item.text(0))
                        else:
                            pass
                        
    def user_deselect_td(self):
        top_item = self.list_artist.topLevelItemCount()
        check_artist = []
        for top_index in range(top_item):
            top_item = self.list_artist.topLevelItem(top_index)
            if top_item.text(0) == 'TD':
                if top_item.checkState(0) == Qt.Checked:
                    top_item.setCheckState(0,Qt.Unchecked)
                else:
                    pass
                top_item.setExpanded(0)
                td_count = self.list_artist.topLevelItem(top_index).childCount()
                for td_index in range(td_count):
                    td_item = self.list_artist.topLevelItem(top_index).child(td_index)
                    if td_item.isDisabled() == False:
                        if td_item.checkState(0) == Qt.Checked:
                            td_item.setCheckState(0,Qt.Unchecked)
                            check_artist.append(td_item.text(0))
                    else:
                        pass
        
    def search_path(self):
        path = QFileDialog.getSaveFileName(filter='xls(*.xls)')[0]
        if path != "":
            self.line_export_path.setText(path)
        else:
            pass
        
    def get_manday_info(self):
        info = Man_day().excel_Signal.connect(self.excel_export)
        
    def excel_export(self):
        path = self.line_export_path.text()
        
        workbook = xlwt.Workbook(encoding='utf-8',style_compression=0)
        alignment = xlwt.Alignment()
        alignment.horz = xlwt.Alignment.HORZ_CENTER
        alignment.vert = xlwt.Alignment.VERT_CENTER
        style = xlwt.XFStyle()
        style.alignment = alignment
        
        check_group = []
        prod_art = []
        fx_art = []
        comp_art = []
        anim_art = []
        td_art = []
        checkstate_user = []
        
        
        if self.checkbox_shot.isChecked():
            project_list = []
            data = []
                
            project_count = self.list_project.count()
            for project_index in range(project_count):
                check_project = self.list_project.item(project_index)
                if check_project.checkState() == Qt.Checked:
                    project_list.append(check_project.text())
                else:
                    pass
                
            for data_index in self.data_info:
                if data_index.split('-')[1] in project_list:
                    data.append(data_index)
                else:
                    pass
                
            filter_data = list(set(data))
            
            for group_index in range(self.list_artist.topLevelItemCount()):
                group_item = self.list_artist.topLevelItem(group_index)
                if group_item.checkState(0) == Qt.Checked or group_item.checkState(0) == Qt.PartiallyChecked:
                    check_group.append(group_item.text(0))
                    for artist_index in range(group_item.childCount()):
                        artist_item = group_item.child(artist_index)
                        if artist_item.isDisabled() == False:
                            if artist_item.checkState(0) == Qt.Checked:
                                if group_item.text(0) == 'Prod':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            prod_art.append(art_info)
                                elif group_item.text(0) == 'FX':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            fx_art.append(art_info)
                                elif group_item.text(0) == 'Comp':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            comp_art.append(art_info)
                                elif group_item.text(0) == 'Anim':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            anim_art.append(art_info)
                                elif group_item.text(0) == 'TD':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            td_art.append(art_info)
                            else:
                                pass

                    if group_item.text(0) == 'Prod':
                        checkstate_user.append(prod_art)
                    if group_item.text(0) == 'FX':
                        checkstate_user.append(fx_art)
                    if group_item.text(0) == 'Comp':
                        checkstate_user.append(comp_art)
                    if group_item.text(0) == 'Anim':
                        checkstate_user.append(anim_art)
                    if group_item.text(0) == 'FX':
                        checkstate_user.append(td_art)
                        
            check_info = dict(zip(check_group,checkstate_user))
            
            for group in check_info:
                sheet = workbook.add_sheet(group,cell_overwrite_ok=False)
                sheet_header = ['工资','扣除税点']
                sheet.write(0,0,'项目',style)
                sheet.col(0).width = 256 * 20
                sheet.write(0,1,'场次',style)
                sheet.col(1).width = 256 * 20
                sheet.write(0,2,'姓名',style)
                sheet.col(2).width = 256 * 10
                sheet.write(0,3,'任务名称',style)
                sheet.col(3).width = 256 * 25
                sheet.write(0,4,'日薪',style)
                sheet.col(4).width = 256 * 10
                sheet.write(0,5,'项目总人天数',style)
                sheet.col(5).width = 256 * 20
                if self.checkbox_dec.isChecked():
                    sheet_header.insert(0,'12月人天数')
                if self.checkbox_nov.isChecked():
                    sheet_header.insert(0,'11月人天数')
                if self.checkbox_oct.isChecked():
                    sheet_header.insert(0,'10月人天数')
                if self.checkbox_sept.isChecked():
                    sheet_header.insert(0,'9月人天数')
                if self.checkbox_aug.isChecked():
                    sheet_header.insert(0,'8月人天数')
                if self.checkbox_jul.isChecked():
                    sheet_header.insert(0,'7月人天数')
                if self.checkbox_jun.isChecked():
                    sheet_header.insert(0,'6月人天数')
                if self.checkbox_may.isChecked():
                    sheet_header.insert(0,'5月人天数')
                if self.checkbox_apr.isChecked():
                    sheet_header.insert(0,'4月人天数')
                if self.checkbox_mar.isChecked():
                    sheet_header.insert(0,'3月人天数')
                if self.checkbox_feb.isChecked():
                    sheet_header.insert(0,'2月人天数')
                if self.checkbox_jan.isChecked():
                    sheet_header.insert(0,'1月人天数')
                
                for header_index in range(len(sheet_header)):
                    col = sheet.col(header_index+6)
                    col.width = 256 * 20
                    sheet.write(0,header_index+6,sheet_header[header_index],style)
                    
                for i in range(len(check_info[group])):
                    project_name = check_info[group][i].split('-')[1]
                    artist_name = check_info[group][i].split('-')[0]
                    seq_name = check_info[group][i].split('-')[2]
                    man_day_total = check_info[group][i].split('-')[3]
                    man_day_jan = check_info[group][i].split('-')[4]
                    man_day_feb = check_info[group][i].split('-')[5]
                    man_day_mar = check_info[group][i].split('-')[6]
                    man_day_apr = check_info[group][i].split('-')[7]
                    man_day_may = check_info[group][i].split('-')[8]
                    man_day_jun = check_info[group][i].split('-')[9]
                    man_day_jul = check_info[group][i].split('-')[10]
                    man_day_aug = check_info[group][i].split('-')[11]
                    man_day_sept = check_info[group][i].split('-')[12]
                    man_day_oct = check_info[group][i].split('-')[13]
                    man_day_nov = check_info[group][i].split('-')[14]
                    man_day_dec = check_info[group][i].split('-')[15]
                    
                    shot_info = eval(check_info[group][i].split('|')[-1])
                    shot_name = shot_info[0]
                    shot_manday = shot_info[1]
                    shot_jan = shot_info[2]
                    shot_feb = shot_info[3]
                    shot_mar = shot_info[4]
                    shot_apr = shot_info[5]
                    shot_may = shot_info[6]
                    shot_jun = shot_info[7]
                    shot_jul = shot_info[8]
                    shot_aug = shot_info[9]
                    shot_sept = shot_info[10]
                    shot_oct = shot_info[11]
                    shot_nov = shot_info[12]
                    shot_dec = shot_info[13]
                    
                    sheet.write(i+1,0,project_name,style)
                    sheet.write(i+1,1,seq_name,style)
                    sheet.write(i+1,2,artist_name,style)
                    sheet.write(i+1,3,shot_name,style)
                    
                    if shot_manday != '':
                        sheet.write(i+1,5,float(shot_manday),style)
                    else:
                        sheet.write(i+1,5,shot_manday,style)
                    if self.checkbox_jan.isChecked():
                        jan_index = sheet_header.index('1月人天数')
                        if shot_jan != '':
                            sheet.write(i+1,6+jan_index,float(shot_jan),style)
                        else:
                            sheet.write(i+1,6+jan_index,shot_jan,style)
                    else:
                        pass
                    if self.checkbox_feb.isChecked():
                        feb_index = sheet_header.index('2月人天数')
                        if shot_feb != '':
                            sheet.write(i+1,6+feb_index,float(shot_feb),style)
                        else:
                            sheet.write(i+1,6+feb_index,shot_feb,style)
                    else:
                        pass
                    
                    if self.checkbox_mar.isChecked():
                        mar_index = sheet_header.index('3月人天数')
                        if shot_mar != '':
                            sheet.write(i+1,6+mar_index,float(shot_mar),style)
                        else:
                            sheet.write(i+1,6+mar_index,shot_mar,style)
                    else:
                        pass
                    
                    if self.checkbox_apr.isChecked():
                        apr_index = sheet_header.index('4月人天数')
                        if shot_apr != '':
                            sheet.write(i+1,6+apr_index,float(shot_apr),style)
                        else:
                            sheet.write(i+1,6+apr_index,shot_apr,style)
                    else:
                        pass
                    
                    if self.checkbox_may.isChecked():
                        may_index = sheet_header.index('5月人天数')
                        if shot_may != '':
                            sheet.write(i+1,6+may_index,float(shot_may),style)
                        else:
                            sheet.write(i+1,6+may_index,shot_may,style)
                    else:
                        pass
                    
                    if self.checkbox_jun.isChecked():
                        jun_index = sheet_header.index('6月人天数')
                        if shot_jun != '':
                            sheet.write(i+1,6+jun_index,float(shot_jun),style)
                        else:
                            sheet.write(i+1,6+jun_index,shot_jun,style)
                    else:
                        pass
                    
                    if self.checkbox_jul.isChecked():
                        jul_index = sheet_header.index('7月人天数')
                        if shot_jul != '':
                            sheet.write(i+1,6+jul_index,float(shot_jul),style)
                        else:
                            sheet.write(i+1,6+jul_index,shot_jul,style)
                    else:
                        pass
                    
                    if self.checkbox_aug.isChecked():
                        aug_index = sheet_header.index('8月人天数')
                        if shot_aug != '':
                            sheet.write(i+1,6+aug_index,float(shot_aug),style)
                        else:
                            sheet.write(i+1,6+aug_index,shot_aug,style)
                    else:
                        pass
                    
                    if self.checkbox_sept.isChecked():
                        sept_index = sheet_header.index('9月人天数')
                        if shot_sept != '':
                            sheet.write(i+1,6+sept_index,float(shot_sept),style)
                        else:
                            sheet.write(i+1,6+sept_index,shot_sept,style)
                    else:
                        pass
                    
                    if self.checkbox_oct.isChecked():
                        oct_index = sheet_header.index('10月人天数')
                        if shot_oct != '':
                            sheet.write(i+1,6+oct_index,float(shot_oct),style)
                        else:
                            sheet.write(i+1,6+oct_index,shot_oct,style)
                    else:
                        pass
                    
                    if self.checkbox_nov.isChecked():
                        nov_index = sheet_header.index('11月人天数')
                        if shot_nov != '':
                            sheet.write(i+1,6+nov_index,float(shot_nov),style)
                        else:
                            sheet.write(i+1,6+nov_index,shot_nov,style)
                    else:
                        pass
                    
                    if self.checkbox_dec.isChecked():
                        dec_index = sheet_header.index('12月人天数')
                        if shot_dec != '':
                            sheet.write(i+1,6+dec_index,float(shot_dec),style)
                        else:
                            sheet.write(i+1,6+dec_index,shot_dec,style)
                    else:
                        pass
                    # workbook.save(path)
        else:
            project_list = []
            data = []
                
            project_count = self.list_project.count()
            for project_index in range(project_count):
                check_project = self.list_project.item(project_index)
                if check_project.checkState() == Qt.Checked:
                    project_list.append(check_project.text())
                else:
                    pass
                
            for data_index in self.data_info:
                if data_index.split('|')[0].split('-')[1] in project_list:
                    data.append(data_index.split('|')[0])
                else:
                    pass
                
            filter_data = list(set(data))
            
            for group_index in range(self.list_artist.topLevelItemCount()):
                group_item = self.list_artist.topLevelItem(group_index)
                if group_item.checkState(0) == Qt.Checked or group_item.checkState(0) == Qt.PartiallyChecked:
                    check_group.append(group_item.text(0))
                    for artist_index in range(group_item.childCount()):
                        artist_item = group_item.child(artist_index)
                        if artist_item.isDisabled() == False:
                            if artist_item.checkState(0) == Qt.Checked:
                                if group_item.text(0) == 'Prod':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            prod_art.append(art_info)
                                elif group_item.text(0) == 'FX':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            fx_art.append(art_info)
                                elif group_item.text(0) == 'Comp':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            comp_art.append(art_info)
                                elif group_item.text(0) == 'Anim':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            anim_art.append(art_info)
                                elif group_item.text(0) == 'TD':
                                    for art_info in filter_data:
                                        if artist_item.text(0) in art_info:
                                            td_art.append(art_info)
                            else:
                                pass

                    if group_item.text(0) == 'Prod':
                        checkstate_user.append(prod_art)
                    if group_item.text(0) == 'FX':
                        checkstate_user.append(fx_art)
                    if group_item.text(0) == 'Comp':
                        checkstate_user.append(comp_art)
                    if group_item.text(0) == 'Anim':
                        checkstate_user.append(anim_art)
                    if group_item.text(0) == 'FX':
                        checkstate_user.append(td_art)
                        
            check_info = dict(zip(check_group,checkstate_user))
            
            for group in check_info:
                sheet = workbook.add_sheet(group,cell_overwrite_ok=False)
                sheet_header = ['工资','扣除税点']
                sheet.write(0,0,'项目',style)
                sheet.col(0).width = 256 * 20
                sheet.write(0,1,'场次',style)
                sheet.col(1).width = 256 * 20
                sheet.write(0,2,'姓名',style)
                sheet.col(2).width = 256 * 10
                sheet.write(0,3,'日薪',style)
                sheet.col(3).width = 256 * 10
                sheet.write(0,4,'项目总人天数',style)
                sheet.col(4).width = 256 * 20
                if self.checkbox_dec.isChecked():
                    sheet_header.insert(0,'12月人天数')
                if self.checkbox_nov.isChecked():
                    sheet_header.insert(0,'11月人天数')
                if self.checkbox_oct.isChecked():
                    sheet_header.insert(0,'10月人天数')
                if self.checkbox_sept.isChecked():
                    sheet_header.insert(0,'9月人天数')
                if self.checkbox_aug.isChecked():
                    sheet_header.insert(0,'8月人天数')
                if self.checkbox_jul.isChecked():
                    sheet_header.insert(0,'7月人天数')
                if self.checkbox_jun.isChecked():
                    sheet_header.insert(0,'6月人天数')
                if self.checkbox_may.isChecked():
                    sheet_header.insert(0,'5月人天数')
                if self.checkbox_apr.isChecked():
                    sheet_header.insert(0,'4月人天数')
                if self.checkbox_mar.isChecked():
                    sheet_header.insert(0,'3月人天数')
                if self.checkbox_feb.isChecked():
                    sheet_header.insert(0,'2月人天数')
                if self.checkbox_jan.isChecked():
                    sheet_header.insert(0,'1月人天数')
                    
                for header_index in range(len(sheet_header)):
                    col = sheet.col(header_index+5)
                    col.width = 256 * 20
                    sheet.write(0,header_index+5,sheet_header[header_index],style)
                    
                for i in range(len(check_info[group])):
                    project_name = check_info[group][i].split('-')[1]
                    artist_name = check_info[group][i].split('-')[0]
                    seq_name = check_info[group][i].split('-')[2] 
                    man_day_total = check_info[group][i].split('-')[3]
                    man_day_jan = check_info[group][i].split('-')[4]
                    man_day_feb = check_info[group][i].split('-')[5]
                    man_day_mar = check_info[group][i].split('-')[6]
                    man_day_apr = check_info[group][i].split('-')[7]
                    man_day_may = check_info[group][i].split('-')[8]
                    man_day_jun = check_info[group][i].split('-')[9]
                    man_day_jul = check_info[group][i].split('-')[10]
                    man_day_aug = check_info[group][i].split('-')[11]
                    man_day_sept = check_info[group][i].split('-')[12]
                    man_day_oct = check_info[group][i].split('-')[13]
                    man_day_nov = check_info[group][i].split('-')[14]
                    man_day_dec = check_info[group][i].split('-')[15]
                    
                    sheet.write(i+1,0,project_name,style)
                    sheet.write(i+1,2,artist_name,style)
                    sheet.write(i+1,1,seq_name,style)
                    
                    if man_day_total != '':
                        sheet.write(i+1,4,float(man_day_total),style)
                    else:
                        sheet.write(i+1,4,man_day_total,style)
                    if self.checkbox_jan.isChecked():
                        jan_index = sheet_header.index('1月人天数')
                        if man_day_jan != '':
                            sheet.write(i+1,5+jan_index,float(man_day_jan),style)
                        else:
                            sheet.write(i+1,5+jan_index,man_day_jan,style)
                    else:
                        pass
                    if self.checkbox_feb.isChecked():
                        feb_index = sheet_header.index('2月人天数')
                        if man_day_feb != '':
                            sheet.write(i+1,5+feb_index,float(man_day_feb),style)
                        else:
                            sheet.write(i+1,5+feb_index,man_day_feb,style)
                    else:
                        pass
                    
                    if self.checkbox_mar.isChecked():
                        mar_index = sheet_header.index('3月人天数')
                        if man_day_mar != '':
                            sheet.write(i+1,5+mar_index,float(man_day_mar),style)
                        else:
                            sheet.write(i+1,5+mar_index,man_day_mar,style)
                    else:
                        pass
                    
                    if self.checkbox_apr.isChecked():
                        apr_index = sheet_header.index('4月人天数')
                        if man_day_apr != '':
                            sheet.write(i+1,5+apr_index,float(man_day_apr),style)
                        else:
                            sheet.write(i+1,5+apr_index,man_day_apr,style)
                    else:
                        pass
                    
                    if self.checkbox_may.isChecked():
                        may_index = sheet_header.index('5月人天数')
                        if man_day_may != '':
                            sheet.write(i+1,5+may_index,float(man_day_may),style)
                        else:
                            sheet.write(i+1,5+may_index,man_day_may,style)
                    else:
                        pass
                    
                    if self.checkbox_jun.isChecked():
                        jun_index = sheet_header.index('6月人天数')
                        if man_day_jun != '':
                            sheet.write(i+1,5+jun_index,float(man_day_jun),style)
                        else:
                            sheet.write(i+1,5+jun_index,man_day_jun,style)
                    else:
                        pass
                    
                    if self.checkbox_jul.isChecked():
                        jul_index = sheet_header.index('7月人天数')
                        if man_day_jun != '':
                            sheet.write(i+1,5+jul_index,float(man_day_jul),style)
                        else:
                            sheet.write(i+1,5+jul_index,man_day_jul,style)
                    else:
                        pass
                    
                    if self.checkbox_aug.isChecked():
                        aug_index = sheet_header.index('8月人天数')
                        if man_day_aug != '':
                            sheet.write(i+1,5+aug_index,float(man_day_aug),style)
                        else:
                            sheet.write(i+1,5+aug_index,man_day_aug,style)
                    else:
                        pass
                    
                    if self.checkbox_sept.isChecked():
                        sept_index = sheet_header.index('9月人天数')
                        if man_day_sept != '':
                            sheet.write(i+1,5+sept_index,float(man_day_sept),style)
                        else:
                            sheet.write(i+1,5+sept_index,man_day_sept,style)
                    else:
                        pass
                    
                    if self.checkbox_oct.isChecked():
                        oct_index = sheet_header.index('10月人天数')
                        if man_day_oct != '':
                            sheet.write(i+1,5+oct_index,float(man_day_oct),style)
                        else:
                            sheet.write(i+1,5+oct_index,man_day_oct,style)
                    else:
                        pass
                    
                    if self.checkbox_nov.isChecked():
                        nov_index = sheet_header.index('11月人天数')
                        if man_day_nov != '':
                            sheet.write(i+1,5+nov_index,float(man_day_nov),style)
                        else:
                            sheet.write(i+1,5+nov_index,man_day_nov,style)
                    else:
                        pass
                    
                    if self.checkbox_dec.isChecked():
                        dec_index = sheet_header.index('12月人天数')
                        if man_day_dec != '':
                            sheet.write(i+1,5+dec_index,float(man_day_dec),style)
                        else:
                            sheet.write(i+1,5+dec_index,man_day_dec,style)
                    else:
                        pass
    
        workbook.save(path)
        export_excel.close(self)
        
class custom_comboBox(QComboBox):
    
    itemChecked = Signal(list)
    check_item_list = []
    
    def __init__(self):
        super().__init__()
        list_widget = QListWidget()
        self.setView(list_widget)
        self.setModel(list_widget.model())
        lineEdit = QLineEdit()
        lineEdit.setReadOnly(True)
        self.setLineEdit(lineEdit)
    
    def add_item(self,text:str):
        checkbox = QCheckBox(text,self.view())
        checkbox.stateChanged.connect(self.check_item)
        self.check_item_list.append(checkbox)
        item = QListWidgetItem(self.view())
        item.setFlags(item.flags())
        self.view().addItem(item)
        self.view().setItemWidget(item,checkbox)
        
    def add_items(self,texts:list):
        texts.insert(0,'All')
        for text in texts:
            self.add_item(text)
        self.check_item_list[0].setChecked(1)
            
    def get_selected(self):
        sel_data = []
        for check in self.check_item_list:
            if self.check_item_list[0] == check:
                continue
            if check.isChecked():
                sel_data.append(check.text())
                
        return sel_data
            
    def select_all_item(self,state):
        check_list = []
        for check in self.check_item_list:
            check.blockSignals(True)
            check.setCheckState(Qt.CheckState(state))
            check.blockSignals(False)
    
    def check_item(self,state):
        check_state = []
        if self.sender() == self.check_item_list[0]:
            self.select_all_item(state)
        else:
            count = len(self.check_item_list)
            for check_index in range(1,count):
                if self.check_item_list[check_index].isChecked():
                    check_state.append(self.check_item_list[check_index])
                else:
                    pass
            self.check_item_list[0].setChecked(0)
            for index in check_state:
                index.setChecked(1)
                
        sel_data = self.get_selected()
        self.itemChecked.emit(sel_data)
        self.lineEdit().setText(','.join(sel_data))

class artist_setting(QMainWindow):
    
    def __init__(self):
        super(artist_setting,self).__init__()
        self.resize(600,600)
        self.setWindowTitle(u'Artist Setting')
        self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        
        self.lab_artist = QLabel('Artist')
        self.lab_artist.setAlignment(Qt.AlignCenter)
        self.list_artist = QTreeWidget()
        self.list_artist.setHeaderHidden(1)
        self.btn_add = QToolButton()
        self.btn_add.setText('+')
        self.btn_remove = QToolButton()
        self.btn_remove.setText('-')
        self.lab_department = QLabel('Department')
        self.combox_department = QComboBox()
        department = ['Prod','FX','Comp','Anim','TD']
        self.combox_department.addItems(department)
        self.lab_CN_Name = QLabel('CN Name')
        self.line_CN_name = QLineEdit('')
        self.lab_name = QLabel('Name')
        self.line_name = QLineEdit('')
        self.rad_enable = QRadioButton('Enabled')
        self.rad_enable.setChecked(1)
        self.rad_disable = QRadioButton('Disabled')
        self.btn_confirm = QPushButton('Confirm')
        
        self.tool_lay = QHBoxLayout()
        self.tool_lay.addWidget(self.btn_add)
        self.tool_lay.addWidget(self.btn_remove)
        
        self.artist_lay = QVBoxLayout()
        self.artist_lay.addWidget(self.lab_artist)
        self.artist_lay.addWidget(self.list_artist)
        self.artist_lay.addLayout(self.tool_lay)
        
        self.department_lay = QFormLayout()
        self.department_lay.setWidget(0,QFormLayout.LabelRole,self.lab_department)
        self.department_lay.setWidget(0,QFormLayout.FieldRole,self.combox_department)
        
        self.name_lay = QHBoxLayout()
        self.name_lay.addWidget(self.lab_name)
        self.name_lay.addWidget(self.line_name)
        
        self.cn_name_lay = QHBoxLayout()
        self.cn_name_lay.addWidget(self.lab_CN_Name)
        self.cn_name_lay.addWidget(self.line_CN_name)
        
        self.check_lay = QHBoxLayout()
        self.check_lay.addWidget(self.rad_enable)
        self.check_lay.addWidget(self.rad_disable)
        
        self.set_lay = QVBoxLayout()
        self.set_lay.addLayout(self.department_lay)
        self.set_lay.addLayout(self.name_lay)
        self.set_lay.addLayout(self.cn_name_lay)
        self.set_lay.addLayout(self.check_lay)
        self.set_lay.addWidget(self.btn_confirm)
        self.set_lay.setContentsMargins(-1, 20, -1, 500)
        
        self.setting_lay = QGridLayout()
        self.setting_lay.addLayout(self.artist_lay,0,0,1,1)
        self.setting_lay.addLayout(self.set_lay,0,1,1,1)
        
        self.childWindow = QWidget()
        self.childWindow.setLayout(self.setting_lay)
        self.setCentralWidget(self.childWindow)
        
        self.get_artists()
        
        self.list_artist.clicked.connect(self.get_artist_setting)
        self.btn_add.clicked.connect(self.add_artist)
        self.btn_remove.clicked.connect(self.remove_artist)
        self.btn_confirm.clicked.connect(self.update_artist)
        self.rad_enable.clicked.connect(self.get_changed)
        self.rad_disable.clicked.connect(self.get_changed)
        
    def get_artists(self):
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        try:
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.users = json.load(f)
        except:
            return 'user.json is not exist!'
        
        for group in self.users:
            group_item = QTreeWidgetItem(self.list_artist)
            group_item.setText(0,group)
            group_item.setFlags(group_item.flags() & ~Qt.ItemIsSelectable)
            sort_group = self.sort_dict(self.users[group])
            for artist in sort_group:
                artist_item = QTreeWidgetItem(group_item)
                artist_item.setText(0,sort_group[artist]['cn_name'])
                
    def sort_dict(self,dict_info):
        reslut = {}
        for i in sorted([i for i in dict_info.keys()]):
            v = dict_info.get(i)
            if isinstance(v,dict):
                v = self.sort_dict(v)
            reslut[i] = v
            
        return reslut
    
    def get_artist_setting(self):
        current_select = self.list_artist.selectedItems()
        
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        try:
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.users = json.load(f)
        
            for index in current_select:
                index_department = index.parent().text(0)
                select_artist = index.text(0)
                for name in self.users[index_department]:
                    if select_artist == self.users[index_department][name]['cn_name']:
                        self.line_name.setText(name)
                        check_stage = self.users[index_department][name]['check']
                        if check_stage == 1:
                            self.rad_enable.setChecked(1)
                            self.rad_disable.setChecked(0)
                        else:
                            self.rad_enable.setChecked(0)
                            self.rad_disable.setChecked(1)
                self.combox_department.setCurrentText(index_department)
                self.line_CN_name.setText(select_artist)
        except:
            return 'user.json is not exist!'
        
    def add_artist(self):
        department = self.combox_department.currentText()
        name = self.line_name.text()
        cn_name = self.line_CN_name.text()
        if self.rad_enable.isChecked():
            check = 1
        if self.rad_disable.isChecked():
            check = 0
            
        new_artist = {
                        name:{
                            'cn_name':cn_name,
                            'check':check
                        }
        }
            
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        try:
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.users = json.load(f)
                
            self.users[department].update(new_artist)
            
            with open(json_path,'w+',encoding='utf-8') as f2:
                json.dump(self.users,f2,indent=4,ensure_ascii=False)
            
            self.list_artist.clear()
            self.get_artists()
        except:
            return 'user.json is not exist!'
        
    def remove_artist(self):
        current_select = self.list_artist.selectedItems()
        
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        try:
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.users = json.load(f)
        
            if current_select:
                for index in current_select:
                    index_department = index.parent().text(0)
                    select_artist = index.text(0)
                    for name in self.users[index_department]:
                        if select_artist == self.users[index_department][name]['cn_name']:
                            source_json = self.users[index_department]
                            del source_json[name]
                                   
                            with open(json_path,'w+',encoding='utf-8') as f2:
                                json.dump(self.users,f2,indent=4,ensure_ascii=False)
                            
                            self.list_artist.clear()
                            self.get_artists()
        except:
            return 'user.json is not exist!'
        
    def get_changed(self):
        if self.rad_enable.isChecked():
            check = 1
        if self.rad_disable.isChecked():
            check = 0
            
        return check
        
    def update_artist(self):
        current_select = self.list_artist.selectedItems()
        
        current_file = os.path.abspath(__file__)
        json_root = os.path.dirname(current_file)
        
        try:
            json_path = '%s/user.json'%(json_root)
            with open(json_path,'r+',encoding='utf-8') as f:
                self.users = json.load(f)
                
            department = self.combox_department.currentText()
            name = self.line_name.text()
            cn_name = self.line_CN_name.text()
            check = self.get_changed()
            

            for index in current_select:
                index_department = index.parent().text(0)
                select_artist = index.text(0)
                for name in self.users[index_department]:
                    if select_artist == self.users[index_department][name]['cn_name']:
                        self.users[department][name]['cn_name'] = str(cn_name)
                        self.users[department][name]['check'] = check
                        
                        with open(json_path,'w+',encoding='utf-8') as f2:
                            json.dump(self.users,f2,indent=4,ensure_ascii=False)
                        
                        self.list_artist.clear()
                        self.get_artists()
                
        except:
            return 'user.json is not exist!'
        
if __name__=='__main__':
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QWidget()
    window = Man_day()
    window.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
    