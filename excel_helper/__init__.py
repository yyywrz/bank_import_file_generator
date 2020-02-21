import xlrd
import path_helper

jan_raw = 'basic_info\\raw\\1\\'
jan_raw_path = path_helper.parent_path(__file__) + '\\' + jan_raw

def notEmpty(array):
    for item in array:
        if item != '':
            return True
    return False


def getData(path):
    data = xlrd.open_workbook(path)
    sheet = data.sheets()[0]
    table = []
    flag = False
    for i in range(0,sheet.nrows):
        row = sheet.row_values(i, start_colx=0, end_colx=None)
        if notEmpty(row):
            if flag:
                table.append(row)
                if '合计' in row : flag = False
            else:
                if ('姓名' in row) or ('实发工资' in row):
                    table.append(row)
                    flag = True
    return(table)

def getmapping(path):
    table = getData(path)
    mapping={}
    for key in table[0]:
        if key !='':
            mapping[key] = table[0].index(key)
    return mapping

def data_dict(path):
    table = getData(path)
    data=[]
    mapping=getmapping(path)
    for instance in table[1:]:
        one={}
        for key,index in mapping.items():
            one[key.replace(' ','')] = str(instance[index]).replace(' ','')
        if (one['姓名'] != '') and (one['实发工资'] != ''):
            data.append(one)
        elif (one['序号']=='合计'):
            modified_one = one
            modified_one['序号'] = ''
            modified_one['姓名'] = '合计'
            data.append(modified_one)

    data_dict = {}
    for instance in data:
        data_dict[instance['姓名']] = instance
    return data_dict

def export_data(path):
    table = data_dict(path)
    data ={}
    for instance in table.values():
        data[instance['姓名']]={
            'name':instance['姓名'],
            'wage':instance['实发工资']
        }
    return data 

def table(path):
    keys=[]

    for key in getmapping(path):
        keys.append(key)
    
    table=[]
    i=0
    for instance in data_dict(path).values():
        table.append([])
        for key in keys:
            table[i].append(instance[key])
        i+=1
    return keys,table


if __name__=="__main__":
    print(getmapping(jan_raw_path+'\\'+'昆阳街办扑火队-工资表1月.xls'))
    print(getData(jan_raw_path+'\\'+'昆阳街办扑火队-工资表1月.xls'))
    print(data_dict(jan_raw_path+'\\'+'昆阳街办扑火队-工资表1月.xls'))
    print(export_data(jan_raw_path+'\\'+'昆阳街办扑火队-工资表1月.xls'))
    print(table(jan_raw_path+'\\'+'昆阳街办扑火队-工资表1月.xls'))