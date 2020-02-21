import path_helper
import file_helper
import export_helper
import os
import time
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, filedialog

people_info_file = 'ppl_info'
people_info_path = path_helper.parent_path(__file__) + '\\' + people_info_file
report_path = path_helper.parent_path(__file__) + '\\导出报告.txt'
export_to_bank_path = path_helper.parent_path(__file__) + '\\银行导出表\\'

def num_vali(n):
    try:
        m=int(n)
        return True
    except Exception as e:
        return False

def all_data(people_info_path):
    all_people_info = {}
    for file_name in os.listdir(people_info_path):
    
        table_name = file_name.split('.')[0]

        all_people_info[ table_name ] = file_helper.getData(people_info_path, file_name)

    return all_people_info

def all_value_for_print():
    all_value=''
    for key,value in all_data(people_info_path).items():
        all_value += key +'\n'
        for personalData in value.values():
            all_value += '\t' + personalData['name'] + ':' + personalData['number'] + '\n'
        all_value += '\n'
    return(all_value)

def table_for_print(name):
    table = file_helper.getData(people_info_path, name + '.data')
    table_data = []
    for value in table.values():
        table_data.append(value['name'] + ': ' + value['number'])
    return table_data

def all_table_names(path):
    table = []
    for file_name in os.listdir(path):
        table_name = file_name.split('.')[0]
        table.append(table_name)
    return(table)

def table_file_related(table, file_name):
    for item in table:
        if item not in file_name:
            return(False)
    return(True)


def centered(win,ww,wh):
    
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw-ww) / 2
    y = (sh-wh) / 2

    win.geometry("%dx%d+%d+%d" %(ww,wh,x,y))


def ppl_info_button():

    def ppl_info_yes_button(*args):
        centered(ppl_info,500,400)
        table_name = table_name_selector.get()
        list_frame = ttk.LabelFrame(ppl_info, text = table_name)
        list_frame.grid(column=0,row=2,padx=40,pady=40)
        ppl_info_list = tk.Listbox(list_frame,width=40)
        list_scrollbar = tk.Scrollbar(list_frame)
        list_scrollbar.pack(side=tk.RIGHT,fill=tk.Y)
        ppl_info_list['yscrollcommand'] = list_scrollbar.set

        for item in table_for_print(table_name):
            ppl_info_list.insert(tk.END,item)

        ppl_info_list.pack(side=tk.LEFT)
        list_scrollbar['command']=ppl_info_list.yview

        def delete_department_button():
            if messagebox.askokcancel('提示', '是否删除'+table_name+'?'):
                file_helper.removeTable(people_info_path, table_name)
                ppl_info.destroy()
                ppl_info_button()

        delete_button = ttk.Button(ppl_info,text='删除该单位', command=delete_department_button)
        delete_button.grid(column=1,row=1)

        operator_frame =  ttk.LabelFrame(ppl_info)
        operator_frame.grid(column=1,row=2)

        def new_one_button():
            def new_one_button_yes():
                name = name_input_box.get()
                number = num_input_box.get()
                if num_vali(number):
                    instance = {
                        'name': name,
                        'number': number
                    }
                    file_helper.addOne(people_info_path,table_name,instance)
                    new_one.destroy()
                    ppl_info_yes_button()
                    
                else:
                    messagebox.showwarning('错误', '检查银行卡号！')

            new_one = tk.Tk()
            centered(new_one,300,70)
            new_one.title('新增个人(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')
        
            ttk.Label(new_one, text='姓名').grid(column=0,row=0)
            name =tk.StringVar()
            name_input_box = ttk.Entry(new_one, width=20, textvariable = name)
            name_input_box.grid(column=1, row=0)

            ttk.Label(new_one, text='银行卡号').grid(column=0,row=1)
            number =tk.StringVar()
            num_input_box = ttk.Entry(new_one, width=20, textvariable = number)
            num_input_box.grid(column=1, row=1)

            yes_button = ttk.Button(new_one,text='确定', command=new_one_button_yes)
            yes_button.grid(column=3,row=0)


        def delete_one_button():
            terms = ppl_info_list.get(ppl_info_list.curselection()).split(':')
            name = terms[0].strip()
            number = terms[1].strip()
            instance = {
                'name': name,
                'number': number
            }
            if messagebox.askokcancel('提示', '是否删除'+table_name+'的'+name+'?'):
                file_helper.deleteOne(people_info_path, table_name, instance)
                ppl_info_yes_button()
            

        def edit_one_button():

            def edit_one_button_yes():
                name = name_input_box.get()
                number = num_input_box.get()
                if num_vali(number):
                    instance = {
                        'name': name,
                        'number': number
                    }
                    if messagebox.askokcancel('提示', '是否更新'+table_name+'的'+name+'?'):
                        file_helper.deleteOne(people_info_path, table_name, old_instance)
                        file_helper.addOne(people_info_path,table_name,instance)
                        edit_one.destroy()
                        ppl_info_yes_button()
                    
                else:
                    messagebox.showwarning('错误', '检查银行卡号！')

            terms = ppl_info_list.get(ppl_info_list.curselection()).split(':')
            old_name = terms[0].strip()
            old_number = terms[1].strip()
            old_instance = {
                'name': old_name,
                'number': old_number
            }

            edit_one = tk.Tk()
            centered(edit_one,300,70)
            edit_one.title('修改个人信息(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')
        
            ttk.Label(edit_one, text='姓名').grid(column=0,row=0)
            name =tk.StringVar()
            name_input_box = ttk.Entry(edit_one, width=20, textvariable = name)
            name_input_box.grid(column=1, row=0)

            ttk.Label(edit_one, text='银行卡号').grid(column=0,row=1)
            number =tk.StringVar()
            num_input_box = ttk.Entry(edit_one, width=20, textvariable = number)
            num_input_box.grid(column=1, row=1)

            yes_button = ttk.Button(edit_one,text='确定', command=edit_one_button_yes)
            yes_button.grid(column=3,row=0)

        def add_many_button():
            def add_many_button_yes():
                number = []
                data = []
                flag = True
                global index
                for i in range(0,index):
                    data.append({
                        'name':names[i].get(),
                        'number':numbers[i].get()})
                    number.append(numbers[i].get())
                    flag = flag and num_vali(numbers[i].get())

                if flag:
                    for instance in data:
                        file_helper.addOne(people_info_path,table_name,instance)
                    many_one.destroy()
                    index=0
                    ppl_info_yes_button()
                    
                else:
                    messagebox.showwarning('错误', '检查银行卡号！')

            many_one = tk.Tk()
            global index
            index = 0
            many_one.title('批量新增(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')
            names = []
            numbers = []
            name_label = []
            number_label = []
            def one_more():
                global index
                name_label.append(ttk.Label(many_one, text='姓名'))
                name_label[index].grid(column=0,row=2*index)

                name = tk.StringVar()
                names.append(ttk.Entry(many_one, width=20, textvariable = name))
                names[index].grid(column=1, row=2*index)

                number_label.append(ttk.Label(many_one, text='银行卡号'))
                number_label[index].grid(column=0,row=index*2+1)

                number =tk.StringVar()
                numbers.append(ttk.Entry(many_one, width=20, textvariable = number))
                numbers[index].grid(column=1, row=index*2+1)
                
                add_button.grid(column=1,row=2*index+2)
                shrink_button.grid(column=2, row=2*index)
                yes_button.grid(column=3,row=2*index+4)

                index+=1
            
            def one_less():
                global index
                if index >0:
                    index-=1

                    names[-1].destroy()
                    del names[-1]
                    numbers[-1].destroy()
                    del numbers[-1]
                    name_label[-1].destroy()
                    del name_label[-1]
                    number_label[-1].destroy()
                    del number_label[-1]
                    add_button.grid(column=1,row=2*index+2)
                    shrink_button.grid(column=2, row=2*index-2)
                    yes_button.grid(column=3,row=2*index+4)
                else:
                    messagebox.showwarning('错误', '没有东西可以删除！')


            add_button = ttk.Button(many_one,text='新增', command=one_more)
            shrink_button = ttk.Button(many_one,text='删除', command=one_less)
            yes_button = ttk.Button(many_one,text='确定', command=add_many_button_yes)
            one_more()


        new_person_button = ttk.Button(operator_frame, text='新增', command=new_one_button)
        new_person_button.grid(column=0,row=0)

        delete_person_button = ttk.Button(operator_frame, text='删除', command=delete_one_button)
        delete_person_button.grid(column=0,row=1)

        edit_person_button = ttk.Button(operator_frame, text='修改', command=edit_one_button)
        edit_person_button.grid(column=0,row=2)

        add_multiple_button = ttk.Button(operator_frame, text='批量新增', command=add_many_button)
        add_multiple_button.grid(column=0,row=4)

    def ppl_info_new_button(*args):

        def new_department_button_yes():
            name = name_input_box.get()
            file_helper.newTable(people_info_path,name)
            new_department.destroy()
            ppl_info.destroy()
            ppl_info_button()

        new_department = tk.Tk()
        centered(new_department,400,100)
        new_department.title('新单位(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')
        
        ttk.Label(new_department, text='输入新单位名称').grid(column=0,row=0, padx=10,pady=10)

        name =tk.StringVar()
        name_input_box = ttk.Entry(new_department, width=20, textvariable = name)
        name_input_box.grid(column=1, row=0)

        yes_button = ttk.Button(new_department,text='确定', command=new_department_button_yes)
        yes_button.grid(column=2,row=0)

    ppl_info = tk.Tk()
    centered(ppl_info,300,150)
    ppl_info.title('人员信息(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')

    table_name_selector = ttk.Combobox(ppl_info,width=20,state='readonly')
    table_name_selector['values'] = all_table_names(people_info_path)
    table_name_selector.grid(column=0,row=0,padx=20,pady=30)
    table_name_selector.current(0)


    yes_button = ttk.Button(ppl_info,text='确定', command=ppl_info_yes_button)
    yes_button.grid(column=1,row=0)

    new_button = ttk.Button(ppl_info,text='增加新单位', command=ppl_info_new_button)
    new_button.grid(column=0,row=1)

def final_export(data):
    suceess_ones={}
    failure_ones={}
    suceess_info=''
    failure_info=''
    for key,value in data.items():
        if value['success']:
            suceess_ones[key] = value
            suceess_info+= key + ': ' + str(len(value['export_data'])) + ' 笔，共计 '+str(value['actual_sum'])+'元。\n'
        else:
            failure_ones[key] = value 
            failure_info+= key + ': ' + '应有'+str(value['expecting_sum']) +'元，已导入'+str(len(value['export_data'])) + ' 笔，共计 '+str(value['actual_sum'])+'元。\n以下人员未导入：'
            for name in value['excel_names']:
                failure_info+=' '+ name
            failure_info+='\n数据库中以下人员不在表中：'
            for name in value['file_names']:
                failure_info+=' '+ name
            failure_info+='\n\n'
    suceess_info+= '\n共计 '+str(len(suceess_ones)) +' 个单位'
    failure_info+= '共计'+str(len(failure_ones)) + ' 个单位'

    final_export_win = tk.Tk()
    suceess_frame = ttk.LabelFrame(final_export_win, text = '成功')
    suceess_frame.grid(column=0,row=0,padx=10,pady=10)
    failure_frame = ttk.LabelFrame(final_export_win, text = '未完成')
    failure_frame.grid(column=1,row=0,padx=10,pady=10)
    suceess_label = tk.Label(suceess_frame, text = suceess_info, justify = tk.LEFT)
    suceess_label.pack()
    failure_label = tk.Label(failure_frame, text = failure_info, justify = tk.LEFT)
    failure_label.pack()

    def export_report():
        with open(report_path,'w') as f:
            f.write('成功:\n'+suceess_info+'\n\n\n未完成:\n'+failure_info)
        final_export_win.destroy()
    
    def export_to_file(key,table):
        with open(export_to_bank_path+key+'.txt','w') as f:
            for terms in table:
                f.write(terms['number']+'|'+terms['wage']+'|'+terms['name']+'|'+terms['other']+'|\n')

    def export_all():
        export_report()
        if os.path.exists(export_to_bank_path):
            for table_file in os.listdir(export_to_bank_path):
                os.remove(export_to_bank_path+'\\'+table_file)
        file_helper.newpath(export_to_bank_path)
        for key,value in data.items():
            export_to_file(key,value['export_data'])
        messagebox.showinfo(title='提示', message='成功')
    
    def report_only():
        export_report()
        messagebox.showinfo(title='提示', message='成功')
    

    def export_success():
        export_report()
        if os.path.exists(export_to_bank_path):
            for table_file in os.listdir(export_to_bank_path):
                os.remove(export_to_bank_path+'\\'+table_file)
        file_helper.newpath(export_to_bank_path)
        for key,value in suceess_ones.items():
            export_to_file(key,value['export_data'])
        messagebox.showinfo(title='提示', message='成功')

    report_only = ttk.Button(final_export_win, text='仅保存报告', command=report_only)
    report_only.grid(column=0,row=1,padx=10,pady=10)
    export_success = ttk.Button(final_export_win, text='保存报告，导出成功表格', command=export_success)
    export_success.grid(column=1,row=1,padx=10,pady=10)
    export_all = ttk.Button(final_export_win, text='保存报告，导出全部', command=export_all)
    export_all.grid(column=2,row=1,padx=10,pady=10)

bank_index = 0
def bank_button_start():
    bank_export = tk.Tk()
    bank_export.title('导入工作表数据(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')    
    global bank_index
    bank_index = 0
    table_name_selector =[]
    path_label = []
    path_button = []

    def get_path_button(label):
        a = filedialog.askopenfilename()
        label.configure(text=a)

    def bank_new_button(table_num=-1,path=''):       
        global bank_index
        table_name_selector.append(ttk.Combobox(bank_export,width=20,state='readonly'))
        table_name_selector[bank_index]['values'] = all_table_names(people_info_path)
        table_name_selector[bank_index].grid(column=0,row=bank_index,padx = 20, pady = 3)
        if table_num>-1:
            table_name_selector[bank_index].current(table_num)
        path_label.append(ttk.Label(bank_export, text=''))
        path_label[bank_index].grid(column=1,row=bank_index)
        current_label = path_label[bank_index]
        if path!='':
            current_label.configure(text=path)
        path_button.append(ttk.Button(bank_export, text='选择文件', command=lambda: get_path_button(current_label)))
        path_button[bank_index].grid(column=2,row=bank_index)
        bank_index+=1
        del_button.grid(column=0,row=bank_index+1)
        new_button.grid(column=1,row=bank_index+1)
        yes_button.grid(column=3,row=bank_index+2)
        directory_import_button.grid(column=1,row=bank_index+2)
        import_count_label.configure(text ='共'+str(bank_index)+'个文件将要被导入')
        import_count_label.grid(column=0,row=bank_index+2)

    new_button = ttk.Button(bank_export, text='增加', command=bank_new_button)
    new_button.grid(column=1,row=bank_index+1)

    def bank_del_button():
        global bank_index
        if bank_index>0:
            bank_index-=1
            table_name_selector[bank_index].destroy()
            del table_name_selector[bank_index]
            path_label[bank_index].destroy()
            del path_label[bank_index]
            path_button[bank_index].destroy()
            del path_button[bank_index]
            del_button.grid(column=0,row=bank_index+1)
            new_button.grid(column=1,row=bank_index+1)
            yes_button.grid(column=3,row=bank_index+2)
            directory_import_button.grid(column=1,row=bank_index+2)
            import_count_label.configure(text ='共'+str(bank_index)+'个文件将要被导入')
            import_count_label.grid(column=0,row=bank_index+2)
        else:
            messagebox.showwarning('错误', '没有东西可以删除！')

    del_button = ttk.Button(bank_export, text='删除', command=bank_del_button)
    del_button.grid(column=0,row=bank_index+1)


    def bank_yes_button():
        data = {}
        global bank_index
        flag = True
        for i in range(0,bank_index):
            if table_name_selector[i].get()=='' or  not (path_label[i]['text'].split('.')[-1] == 'xls' or path_label[i]['text'].split('.')[-1] == 'xlsx'):
                flag = False
        if flag:
            for i in range(0, bank_index):
                (success,
                excel_names,
                file_names,
                export_data,
                expecting_sum,
                actual_sum) = export_helper.experter(
                    path_label[i]['text'], 
                    people_info_path,
                    table_name_selector[i].get())
                
                data[table_name_selector[i].get()] = {
                    'success':success,
                    'excel_names':excel_names,
                    'file_names':file_names,
                    'export_data':export_data,
                    'expecting_sum':expecting_sum,
                    'actual_sum':actual_sum
                }
            final_export(data)
        else:
            messagebox.showwarning('错误', '检查文件路径或对应的人员信息表！')
    yes_button = ttk.Button(bank_export,text='确定', command=bank_yes_button)
    yes_button.grid(column=3,row=bank_index+2,padx=10,pady=10)
    
    def directory_import_button():
        global bank_index
        while bank_index>0:
            bank_del_button()
        table_list = all_table_names(people_info_path)
        file_path = filedialog.askdirectory()
        file_list = []
        for file_name in os.listdir(file_path):
            file_list.append(file_name)
        table_check ={}
        for item in table_list:
            table_check[item] = True
        for one_file in file_list:
            for one_table in table_list:
                if table_file_related(one_table, one_file) and table_check[one_table]:
                    bank_new_button(table_list.index(one_table),str(file_path).replace('/','\\')+'\\'+one_file)
                    table_check[one_table] = False
                    break
        if bank_index<1:
            messagebox.showwarning('错误', '没有任何文件被导入\n检查文件路径！')


    directory_import_button = ttk.Button(bank_export,text ='从文件夹自动导入', command=directory_import_button)
    directory_import_button.grid(column=1,row=bank_index+2,padx=10,pady=10)

    import_count_label = ttk.Label(bank_export,text ='共0个文件将要被导入')
    import_count_label.grid(column=0,row=bank_index+2,padx=10,pady=10)

    bank_new_button()

def main_run():
    
    root = tk.Tk()
    root.title('晋宁合宏劳务(Copyright © 2020 Rongzhen Wei. All Rights Reserved)')

    ttk.Label(root, text='您好！欢迎使用本系统！').grid(column=0,row=0, padx=300,pady=50)

    operation_box =  ttk.LabelFrame(root)
    operation_box.grid(column=0,row=1,pady=50)

    bank_button = ttk.Button(operation_box,text='生成银行导出表', command=bank_button_start)
    bank_button.grid(column=0,row=0,padx=40,pady=10)

    
    people_info_button = ttk.Button(operation_box,text='人员信息', command=ppl_info_button)
    people_info_button.grid(column=0,row=2,padx=40,pady=10)


    centered(root,750,450)

    root.mainloop()



if __name__ == '__main__':
    main_run()