## dependencies
from os import listdir, path

## target directories
# tgtDirs = ['SALES - Sales Channel Revenue', 'SALES - Sales Channel Revenue SF', 'Sales - Special Focus' \
#           ,'SALES - Transaction Analytics', 'SALES - Transaction Analytics SF', 'SALES - Transaction Data']
tgtDirs = ['SALES - Callidus Extract']

## root directory
rootDir = 'C:/Users/RNath000/source/repos/BDI/sales/jobs/'
tgtFiles = [path.join(rootDir, tgtFldrs, instFile) for tgtFldrs in listdir(rootDir) if path.isdir(path.join(rootDir, tgtFldrs)) and tgtFldrs in tgtDirs \
                                                   for instFile in listdir(path.join(rootDir, tgtFldrs)) if instFile == 'install.sql']

## source lists
sourceList = []
intoList = []

## populate into list
for inFile in tgtFiles:
    with open(inFile, 'r') as readFile:
        lines = readFile.readlines()

        for line in lines:
            if 'into' in line.lower():
                intoLine = line.lower().split('into')[1].lstrip().split(' ')[0].replace('\n','').replace(';','').replace(')','').replace('[','').replace(']','').replace(',','').replace('\'','') 

                if '(' in intoLine:
                    intoLine = intoLine.split('(')[0]

                if '#' not in intoLine and intoLine.lower() not in intoList:
                    intoList.append(intoLine.strip(' '))

## populate src list
for inFile in tgtFiles:
    with open(inFile, 'r') as readFile:
        lines = readFile.readlines()

        for line in lines:
            if 'from' in line.lower():
                fromLine = line.lower().split('from')[1].lstrip().split(' ')[0].replace('\n','').replace(';','').replace(')','').replace('[','').replace(']','').replace(',','').replace('\'','') 
                
                if '#' not in fromLine and fromLine.lower() not in sourceList and fromLine.lower() not in intoList:
                    sourceList.append(fromLine.strip(' '))

            if 'join' in line.lower():
                joinLine = line.lower().split('join')[1].lstrip().split(' ')[0].replace('\n','').replace(';','').replace(')','').replace('[','').replace(']','').replace(',','').replace('\'','') 

                if '#' not in joinLine and joinLine.lower() not in sourceList and fromLine.lower() not in intoList:
                    sourceList.append(joinLine.strip(' '))

## sort list and print
sourceList.sort()
intoList.sort()

# print('SourceList: ')

# for src in sourceList:
#     if src.count('.') == 2 and 'msdb' not in src.lower() and 'common' not in src.lower():
#         print(src)


print('\nIn List: ')

for src in intoList:
    if src.count('.') == 2 and 'msdb' not in src.lower() and 'common' not in src.lower() and src not in sourceList:
        print(src)