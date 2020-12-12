#-----------------------------Author- Pratik Dutta--------------#
#-----------------------------App Version 1.0
#-----------------------------Automation using python 3.8.2
#---It will  find the string Only from Word Document with .docx extention file.

from docx2txt import *
import os
from docx import Document
print("*************************Python Automation ***********************\n")
print("It will find string only from .docx file and return the full file location\n")
count=0
filepath=input("Enter Root folder path- e.g-D:\Document\n")
word=input("\nNote:-Search String Case sensitive\nEnter your Search String: \n")
print("\n")
for root, dirs,files in os.walk(filepath):
    for file in files:
        if file.endswith('.docx'):
            path=(os.path.join(root,file))
            res=docx2txt.process(path)
            if word in res:
                print(path)
                count=count+1
            else:
                break;
if count==0:
    print("No matched String found.....!!!!\nPlease run the program again with Different string")
input("\n\nPress enter key to Exit")
            
