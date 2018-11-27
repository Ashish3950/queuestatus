'''
Created on Nov 27, 2018

@author: ashish.maikhuri
'''
import Excle 
import os
from docx  import Document
os.chdir('D:\\test')
doc=Document('test.docx')

for s,i in enumerate(doc.paragraphs,1):
    if(s==5):
        temp=Excle.totalincident()
        i.text=temp
    elif(s==6):
        temp=Excle.AUFincident()
        i.text=temp
    elif(s==7):
        temp=Excle.Assignedincident()
        i.text=temp
    elif(s==8):
        temp=Excle.wipincident()
        i.text=temp
    elif(s==9):
        temp=Excle.A3partyincident()
        i.text=temp
    elif(s==10):
        temp=Excle.Achangeincident()
        i.text=temp
    
        