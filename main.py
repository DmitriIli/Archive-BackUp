
# Export arhive 
'''
<parameters>
	<company>MyNano</company>
	<title>Export video</title>
	<version>1.0</version>
</parameters>
'''
import json
import os
import time
import datetime
import xlrd
from video_exporter import VideoExporter, status
from exthttp import create_app, BaseHandler, http

DIR = 'E:\\2024\\'
CHANNEL_GUID_DICT={"AS2/1":"SAqdbNag_IoorJMih","AS2/5":"qFeFf3lF_IoorJMih","AS2/16":"EbneHxpI_BMZiE99P"}
tasks = {}


def callback(guid, state):
    host.message("[%s] %s" % (guid, state))
    if state == status.success:
        fpath = tasks.pop(guid)
        host.message("[%s] %s  success" % (guid, state))
    elif state == status.error:
        fpath = tasks.pop(guid)
        host.error("Export failed: %s" % fpath)

def get_dirs_list_on_disk():
	try:
		#os.path.exists('E:\\')
		dir = DIR
		os.chdir(dir)
		lst = os.listdir(dir)
		sub_dir_on_disk = [int(name) for name in lst if (os.path.isdir((os.path.join(dir, name))))]
		sub_dir_on_disk.sort(reverse = False)
		sub_dir_on_disk_empty = [int(name) for name in lst if (os.path.isdir((os.path.join(dir, name))) and not os.listdir(os.path.join(dir, name)))]
	except:
		host.error('внешний диск')
	else:
		return(sub_dir_on_disk,sub_dir_on_disk_empty)

def get_weeks_list_on_reestr():
	try:
		dir = DIR
		os.chdir(dir)
		book = xlrd.open_workbook('reestr.xls')
		sh = book.sheet_by_index(0)
		lst = []
		for rx in range(sh.nrows):
			if sh.row(rx)[1].value and sh.row(rx)[0].value.find('-') != -1:
				index = sh.row(rx)[0].value.find('-')
				lst.append(int(sh.row(rx)[0].value[0:index]))
			result = list(set(lst))
			result.sort(reverse = False)
	except:
		host.error('нет файла реестра')
	else:
		return result

def get_archive_dict(weeks):
	lst_archive = dict()
	book = xlrd.open_workbook('reestr.xls')
	sh = book.sheet_by_index(0)
	host.message(lst_archive.items())
	for week in weeks:
		for rx in range(sh.nrows):
			index = sh.row(rx)[0].value.find('-')
			if sh.row(rx)[0].value[0:index] == str(week):
				lst = []
				try:
					lst_archive[week]
				except:
					lst_archive[week] = []
				data = xlrd.xldate.xldate_as_datetime(sh.row(rx)[1].value, 0)
				start = xlrd.xldate.xldate_as_datetime(sh.row(rx)[2].value, 0).time()
				start = datetime.datetime.combine(data, start)
				lst.append(start)
				end = xlrd.xldate.xldate_as_datetime(sh.row(rx)[3].value, 0).time()
				end = datetime.datetime.combine(data, end)
				lst.append(end)
				lst_archive[week].append(lst)
	return lst_archive

#with open('text_{}.json'.format(count), 'w') as f:
#				json.dump(result[week], f)


def add_archive():
	ve = VideoExporter()
	try:
		dirs = get_dirs_list_on_disk()
		weeks = get_weeks_list_on_reestr()
		res = list(set.difference(set(weeks), set(dirs[0])))
		res = res + dirs[1]
		res.sort(reverse = False)
	except:
		host.error('что-то пошло не так')
	else:
		result = get_archive_dict(res)
		for week_lst in result:
			dir = DIR
			os.chdir(dir)
			try:
				os.mkdir(str(week_lst))
			except:
				host.error('папка {} существует'.format(str(week_lst)))
			finally:
				dir = os.path.join(dir,str(week_lst))
				ve = VideoExporter(callback = callback)
				ve.file_name_tmpl = "{name}({dt_start} - {dt_end}){sub}.avi"
				ve.export_folder = str(dir)
				ve.cancel_task_with_states = tuple()
				
				for item in result[week_lst]:
					for guid in CHANNEL_GUID_DICT.values():
						task_guid, file_path = ve.export(guid,item[0],item[1])
						tasks[task_guid] = file_path

host.activate_on_shortcut('F11',add_archive)
