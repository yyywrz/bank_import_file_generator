import path_helper
import os

people_info_file = 'ppl_info'
people_info_path = path_helper.parent_path(__file__) + '\\' + people_info_file
print(people_info_path)
for file_name in os.listdir(people_info_path):    
    print(file_name)
    table_name = file_name.split('.')[0] + '.data'
    try:
        os.rename(people_info_path+'\\'+file_name, people_info_path+'\\'+table_name)
    except Exception as e:
        print(e)
        print('rename file fail')
    else:
        print('rename file success')
