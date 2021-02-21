import shutil, os

AbsolutePath = r'C:\Python\Projects and Scripts\File folder operations' 

counter = 0

wrongname = 'Specification'
correctname = 'Specfication'

for WrongFileName in os.listdir(AbsolutePath):

    print(WrongFileName)
#print(WrongFileName.find(wrongname))
    if WrongFileName.find(wrongname) == -1:
        continue

    counter += 1


    correctFileName = WrongFileName.replace(wrongname,correctname)
    #print(correctFileName)
    
    # construct file location/file name
    wrongFile = AbsolutePath + '\\' + WrongFileName
    correctFile = AbsolutePath + '\\' + correctFileName
    
    # do renaming
    shutil.move(wrongFile, correctFile)

print("files renamed are:", counter)