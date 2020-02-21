import os


def newpath(path):
    if not os.path.exists(path):
        os.makedirs(path)

def newTable(path, name):
    file = path + '\\' + name + '.data'
    if not os.path.exists(file):
        with open(file,'w') as f:
            pass

def removeTable(path, name):
    os.remove(path + '\\' + name + '.data')

def getData(path, name):
    file = path + '\\' + name
    with open(file,'r') as f:
        table_list = {}
        for line in f.readlines():
            terms = line.replace('\n','').split('|')
            if len(terms)>1:
                table_list[terms[1]] = {
                    'name': terms[1],
                    'number': terms[0]
                }
        return(table_list)

def addOne(path, name, instance):
    file = path + '\\' + name + '.data'
    with open(file,'a') as f:
        f.write(str(instance['number'])+'|'+instance['name']+'\n')

def rewrite(path, name, data):
    file = path + '\\' + name + '.data'
    with open(file,'w') as f:
        for instance in data.values():
            f.write(str(instance['number'])+'|'+instance['name']+'\n')

def deleteOne(path, name, one):
    data = getData(path, name + '.data')
    del data[one['name']]
    rewrite(path, name, data)
