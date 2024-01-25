# -*- coding: utf-8 -*-

# import cgtw2

# cgt = cgtw2.tw()
import sys,os
sys.path.append(os.getcwd())

sys.path.append(r"D:\CgTeamWork_v7\bin\base")
sys.path.append(r"C:\CgTeamWork_v7\bin\base")
sys.path.append(r"E:\CgTeamWork_v7\bin\base")
import cgtw2
mcgt = cgtw2.tw()

import datetime

class project:
    
    @staticmethod
    def get_active_project():
        projectId_list = mcgt.info.get_id(db='public',module='project',filter_list=[['project.status','=','Active']])
        project_data = mcgt.info.get(db='public',module='project',id_list=projectId_list,field_sign_list=['project.entity','project.database','project.create_time','project.last_update_time'])
        
        return project_data
    
    @staticmethod
    def get_status_color():
        status_color = mcgt.status.get_status_and_color()
        
        # status_color = mcgt.status.get_status_and_color()
        # status = list(set([k for k in status_color]))
        
        return status_color
    
    def get_active_project_dict():
        active_cgt_project = project.get_active_project()
        
        project_name = []
        project_db = []
        
        for i in active_cgt_project:
            create_time = datetime.datetime.strptime(i['project.create_time'],'%Y-%m-%d %H:%M:%S')
            if str(create_time.year) >= '2023':
                project_name.append(i['project.entity'])
                project_db.append(
                    {
                        i['project.database']:{
                            'project_id':i['id'],
                            'create_time':i['project.create_time'],
                            'update_time':i['project.last_update_time']
                        }
                    }
                )
            
        reslut = dict(zip(project_name,project_db))
        
        return reslut
    
    def filter_project(project_name):
        projects = project.get_active_project_dict()
        proj_db = list(projects[project_name])
        
        return proj_db
    
    def get_seq(db):
        proj_db = 'proj_%s'%(db.lower())
        task_id = mcgt.task.get_id(proj_db,'shot',filter_list=[])
        task = mcgt.task.get(proj_db,'shot',task_id,field_sign_list=['seq.entity'])

        seq = list(set([i['seq.entity'] for i in task]))
        
        return seq
    
    def get_artist(db):
        proj_db = 'proj_%s'%(db.lower())
        task_id = mcgt.task.get_id(proj_db,'shot',filter_list=[])
        task = mcgt.task.get(proj_db,'shot',task_id,field_sign_list=['task.artist'])
        
        artist = list(set([i['task.artist'] for i in task if i['task.artist'] != '']))
        
        return artist
    
class shot:
    def get_project_count(db,seq=None,status_type=None):
        total_count = mcgt.task.count(db,'shot',filter_list=[])
        if seq == None:
            if status_type == None:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[])
            else:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=[status_type],filter_list=[])
        else:
            if status_type == None:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['seq.entity','in',seq]])
            else:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=[status_type],filter_list=[['seq.entity','in',seq]])
        
        if status_type == None:
            project_reslut = [{i['task.status']:(int(i['count'])/int(total_count)*100)} for i in task_count]
        else:
            project_reslut = [{i[status_type]:(int(i['count'])/int(total_count)*100)} for i in task_count]
        
        return project_reslut
        
    def get_task_pipeline(db,seq=None):
        if seq == None:
            task_id = mcgt.task.get_id(db,'shot',filter_list=[])
        else:
            task_id = mcgt.task.get_id(db,'shot',filter_list=[['seq.entity','=',seq]])
            
        filter_sign = ['pipeline.entity']
        get_pipeline = mcgt.task.get(db,'shot',task_id,field_sign_list=filter_sign)
        
        pipeline = list(set([i['pipeline.entity'] for i in get_pipeline]))
        
        return pipeline
    
    def get_task_status(db,seq=None,status=None):
        if seq == None:
            task_id = mcgt.task.get_id(db,'shot',filter_list=[])
        else:
            task_id = mcgt.task.get_id(db,'shot',filter_list=[['seq.entity','=',seq]])
        
        if status == None:
            filter_sign = ['task.status']
            get_status = mcgt.task.get(db,'shot',task_id,field_sign_list=filter_sign)
            
            status = list(set([i['task.status'] for i in get_status]))
        else:
            filter_sign = [status]
            get_status = mcgt.task.get(db,'shot',task_id,field_sign_list=filter_sign)
            
            status = list(set([i[status] for i in get_status]))
        
        return status
    
    def get_task_count(db,seq=None,pipeline=str,status=str):
        if seq == None:
            task_count = mcgt.task.count(db,'shot',filter_list=[['pipeline.entity','=',pipeline],['task.status','=',status]])
        else:
            task_count = mcgt.task.count(db,'shot',filter_list=[['seq.entity','in',seq],['pipeline.entity','=',pipeline],['task.status','=',status]])
        
        return task_count
    
    def get_task_count_by_artist(db,seq=None,artist=str,status=str):
        if seq == None:
            task_count = mcgt.task.count(db,'shot',filter_list=[['task.artist','=',artist],['task.status','=',status]])
        else:
            task_count = mcgt.task.count(db,'shot',filter_list=[['seq.entity','in',seq],['task.artist','=',artist],['task.status','=',status]])
            
        return task_count
    
    def get_task_artist_count(db,seq=None,stauts_type=None,artist=str):
        if seq == None:
            if stauts_type == None:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['task.artist','=',artist]])
            else:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=[stauts_type],filter_list=[['task.artist','=',artist]])
        else:
            if stauts_type == None:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['seq.entity','in',seq],['task.artist','=',artist]])
            else:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=[stauts_type],filter_list=[['seq.entity','in',seq],['task.artist','=',artist]])
            
        # reslut_count = {pipeline:task_count}
        if stauts_type == None:
            count_dict = [{i['task.status']:i['count']} for i in task_count]
        else:
            count_dict = [{i[stauts_type]:i['count']} for i in task_count]
        
        return count_dict
    
    def get_task_pipeline_count(db,seq=None,stauts_type=None,pipeline=str):
        if seq == None:
            if stauts_type == None:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['pipeline.entity','=',pipeline]])
            else:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=[stauts_type],filter_list=[['pipeline.entity','=',pipeline]])
        else:
            if stauts_type == None:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['seq.entity','in',seq],['pipeline.entity','=',pipeline]])
            else:
                task_count = mcgt.task.group_count(db,'shot',field_sign_list=[stauts_type],filter_list=[['seq.entity','in',seq],['pipeline.entity','=',pipeline]])
        
        if stauts_type == None:
            count_dict = [{i['task.status']:i['count']} for i in task_count]
        else:
            count_dict = [{i[stauts_type]:i['count']} for i in task_count]
        
        return count_dict

# proj = 'YMDZX'
# db = 'proj_ymdzx'
# status = shot.get_task_status(db)
# artist = project.get_artist(proj)

# import json
# with open('D:/TD_Project/showDashboard/check_task.json','w') as f:
#     json.dump(test,f,indent=4)

# proj = 'YMDZX'
# db = 'proj_ymdzx'
# task_count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[])
# total_count = mcgt.task.count(db,'shot',filter_list=[])

# test = [{i['task.status']:(int(i['count'])/int(total_count)*100)} for i in task_count]
# test = shot.get_project_count(db)
# print(test)

# db = 'proj_ymdzx'
# seq_test = ['VFX','prod','dev']
# pipeline_test = 'fx'

# task_count = []
# for seq in seq_test:
#     for pipeline in pipeline_test:
#         count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['seq.entity','=',seq],['pipeline.entity','=',pipeline]])
#         if len(count) != 0:
#             task_count.append(count)
#         else:
#             pass
# count = mcgt.task.group_count(db,'shot',field_sign_list=['task.status'],filter_list=[['seq.entity','in',seq_test],['pipeline.entity','in',pipeline_test]])
# print(count)

# count =shot.get_task_artist_count(db,seq=seq_test,stauts_type='task.sup_review',artist='王晨阳')
# print(count)