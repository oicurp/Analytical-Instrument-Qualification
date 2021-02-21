import shutil
import os
import sys

# preparation of files

AbsolutePath = r"C:\Users\305011902\Box Sync\QC Internal Transmit Files\C and Q\050_Installation and Operation Qualification"

counter = 0
wrong_text = 'Quality Control (QC)'
correct_text = 'QC'

# end of file preparation

# loop thru the list for file names to be corrected.

for FileName in os.listdir(AbsolutePath):
        
    if FileName.find(wrong_text) == -1:
        continue
    
    counter +=1
    correctFileName = FileName.replace(wrong_text,correct_text)
    
    # construct file location/file name
    wrongFile = AbsolutePath + '\\' + FileName
    correctFile = AbsolutePath + '\\' + correctFileName
    
    # do renaming
    shutil.move(wrongFile, correctFile)

# Output the number of file names that have been corrected.
print("Total number of files with name corrected is:", counter)
