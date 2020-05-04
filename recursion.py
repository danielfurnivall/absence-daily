import os

startdir = input("What's the start dir (forward slashes only)")


def list_dir(dir, indent_level):
    dirs = os.listdir(dir)
    for i in dirs:
        if os.path.isdir(dir+"/"+i):
            indent_level += 1
            print(("|__"*indent_level)+i)
            try:
                list_dir(dir + "/" + i, indent_level)
            except(PermissionError):
                continue
            indent_level -= 1
        else:
            print((indent_level *"|__")+" "+i)

list_dir(startdir, 0)

