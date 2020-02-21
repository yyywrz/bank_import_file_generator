import os


def current_path(name=__file__):
    return os.path.abspath(name)

def parent_path(name=__file__):
    parent = os.path.abspath(os.path.dirname(current_path(name)) + os.path.sep + ".")
    return parent

def grandparent_path(name=__file__):
    return parent_path(parent_path(name))

if __name__=='__main__':
    c=current_path()
    print(c)
    print(parent_path())
    print(grandparent_path())