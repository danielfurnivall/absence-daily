import os

dir = (os.listdir('C:/gwplogs/gwp_logs'))
with open('C:/gwplogs/gwplogs.txt', 'w', encoding='ANSI') as outfile:
    for file in dir:
        with open('C:/gwplogs/gwp_logs/'+file, 'r', encoding='ANSI', errors='backslashreplace') as infile:
            outfile.write(infile.read())
