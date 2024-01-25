# -*- coding: utf-8 -*-

# import cgtw2

# cgt = cgtw2.tw()
import sys,os
sys.path.append(os.getcwd())
try:
    from cgt_api.mcgt import mcgt
except:
    sys.path.append(r"D:\CgTeamWork_v7\bin\base")
    sys.path.append(r"C:\CgTeamWork_v7\bin\base")
    sys.path.append(r"E:\CgTeamWork_v7\bin\base")
    import cgtw2
    mcgt = cgtw2.tw()
import datetime
from dateutil.relativedelta import relativedelta
    
class cgt_common:
    
    @staticmethod
    def get_current_user():
        account = mcgt.account.get_id(filter_list=[['account.id','=',mcgt.login.account_id()]])
        user_data = mcgt.account.get(account,field_sign_list=['account.entity','account.name','account.group'])
        
        for i in user_data:
            user_name = i['account.entity']
            cn_name = i['account.name']
            user_group = i['account.group']
            
            user = {
                user_name:{
                    'cn_name':cn_name,
                    'group':user_group
                }
            }
            
        return user

class cgt_project:
    
    @staticmethod
    def get_active_project():
        projectId_list = mcgt.info.get_id(db='public',module='project',filter_list=[['project.status','=','Active']])
        project_data = mcgt.info.get(db='public',module='project',id_list=projectId_list,field_sign_list=['project.entity','project.database','project.create_time','project.last_update_time'])
        
        return project_data
        
class cgt_shot:
    
    @staticmethod
    def get_task_info(db):
        account = mcgt.account.get_id(filter_list=[['account.id','=',mcgt.login.account_id()]])
        user_data = mcgt.account.get(account,field_sign_list=['account.entity','account.name','account.group'])
        for i in user_data:
            if '人天工具' not in i['account.group']:
                task_id = mcgt.task.get_id(db,'shot',filter_list=[['task.account','=','%s'%(i['account.entity'])]])
            else:
                task_id = mcgt.task.get_id(db,'shot',filter_list=[])
                
        filter_sign = ['eps.project_code','shot.entity','task.artist','task.mandays','task.md_jan','task.md_feb','task.md_mar','task.md_apr','task.md_may','task.md_jun','task.md_jul','task.md_aug','task.md_sep','task.md_oct','task.md_nov','task.md_dec','task.entity','seq.entity','shot.create_time','shot.last_update_time','pipeline.entity','seq.show_workyear']
        task = mcgt.task.get(db,'shot',task_id,field_sign_list=filter_sign)
        
        return task

class info:
    
    def get_active_project_dict():
        projects = []
        projects_db = []
        project_data = cgt_project.get_active_project()
        for index in project_data:
            project_create_time = index['project.create_time']
            create_time = datetime.datetime.strptime(project_create_time,'%Y-%m-%d %H:%M:%S')
            if str(create_time.year) >= str(2023):
            
                projects.append(index['project.entity'])
                projects_db.append(index['project.database'])
            
        result = dict(zip(projects_db,projects))
            
        return result

    def get_active_projects():
        projects = []
        project_data = cgt_project.get_active_project()
        for index in project_data:
            project_create_time = index['project.create_time']
            create_time = datetime.datetime.strptime(project_create_time,'%Y-%m-%d %H:%M:%S')
            if str(create_time.year) >= str(2023):
                projects.append(index['project.entity'])
            
        return projects
    
    def get_tasks():
        artists = []
        project_db = []
        tasks = []
        project_data = cgt_project.get_active_project()
        for project_index in project_data:
            project_db.append(project_index['project.database'])
            
        for db in project_db:
            tasks_info = cgt_shot.get_task_info(db)
            tasks.append(tasks_info)
            
        return tasks
    
    def get_task_manday(db):
        project_dict = info.get_active_project_dict()
        project = project_dict[db]
        
        task_info = cgt_shot.get_task_info(db)
        result = {}
        artists = []
        shot = []
        test = []
        for task in task_info:
            artist = task['task.artist']
            if artist == "":
                artist = '待分配'
            if len(artist.split(',')) > 1:
                artist = artist.split(',')[0]
                
            task_create_year = task['shot.create_time']
            task_last_updata = task['shot.last_update_time']
            date_create = datetime.datetime.strptime(task_create_year,'%Y-%m-%d %H:%M:%S')
            date_update = datetime.datetime.strptime(task_last_updata,'%Y-%m-%d %H:%M:%S')
            
            artists.append({
                            artist:{
                                task['seq.entity']:{
                                    ('%s-%s')%(task['shot.entity'],task['task.entity']):{
                                        'mandays':task['task.mandays'],
                                        'jan':task['task.md_jan'],
                                        'feb':task['task.md_feb'],
                                        'mar':task['task.md_mar'],
                                        'apr':task['task.md_apr'],
                                        'may':task['task.md_may'],
                                        'jun':task['task.md_jun'],
                                        'jul':task['task.md_jul'],
                                        'aug':task['task.md_aug'],
                                        'sep':task['task.md_sep'],
                                        'oct':task['task.md_oct'],
                                        'nov':task['task.md_nov'],
                                        'dec':task['task.md_dec'],
                                        'create_time':task['shot.create_time'],
                                        'update_time':task['shot.last_update_time'],
                                        'pipline_step':task['pipeline.entity'],
                                        'project_workyear':task['seq.show_workyear']
                                    }
                                }
                            }
                })
            
            result_artist = {}
            result_project = {}
            result = {}
            
            for index in artists:
                for k_artist,v_artist in index.items():
                    result_artist.setdefault(k_artist,[]).append(v_artist)
                    
            for artist_index in result_artist:
                artist_tasks = result_artist[artist_index]
                result_task = {}
                for task in artist_tasks:
                    for k_seq,v_seq in task.items():
                        result_task.setdefault(k_seq,{}).update(v_seq)
                result_artist[artist_index] = result_task
                    
            for item in result_artist:
                result_project[project] = result_artist[item]
                result.setdefault(item,{}).update(result_project)
    
        return result
    
    def get_tasks_manday():
        project_dict = info.get_active_project_dict()
        tasks_info = []
        
        for project_db in project_dict:
            if project_db != 'proj_com23':
                tasks_info.append(info.get_task_manday(project_db))
            
        result = {}
        
        for index in tasks_info:
            for k,v in index.items():
                result.setdefault(k,{}).update(v)
    
        return result
