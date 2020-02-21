import excel_helper
import file_helper
import path_helper
from tkinter import messagebox

jan_raw = 'basic_info\\raw\\1\\'
jan_raw_path = path_helper.parent_path(__file__) + '\\' + jan_raw

people_info_file = 'ppl_info'
people_info_path = path_helper.parent_path(__file__) + '\\' + people_info_file

def experter(excel_path,file_path,name):
    excel_data = excel_helper.export_data(excel_path)
    excel_names=[]
    for key in excel_data:
        excel_names.append(key)

    file_info = file_helper.getData(file_path, name + '.data')
    file_names=[]
    for key in file_info:
        file_names.append(key)
    
    data=[]
    sum = 0
    for key,value in excel_data.items():
        instance = {}
        if key in file_names:
            del excel_names[excel_names.index(key)]
            del file_names[file_names.index(key)]
            instance={
                'number': file_info[key]['number'],
                'wage': str(round(float(value['wage']),2)),
                'name': key,
                'other': '工资'
            }
            sum+=float(value['wage'])
            data.append(instance)
    try:
        expection = excel_data['合计']['wage']
        del excel_names[excel_names.index('合计')]
    except Exception as e:
        expecion = 0
        messagebox.showwarning('错误', '找不到表格中的“合计”！')

    if round(float(sum),2) == round(float(expection),2):
        return(True,excel_names,file_names,data,round(float(expection),2),round(float(sum),2))
    else:
        return(False,excel_names,file_names,data,round(float(expection),2),round(float(sum),2))

if __name__ == '__main__':
    print(experter(jan_raw_path+'\\'+'昆阳街办扑火队-工资表1月.xls',people_info_path,'昆阳街办扑火队'))