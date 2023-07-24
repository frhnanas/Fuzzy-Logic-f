# -*- coding: utf-8 -*-
"""
Created on Sun Nov 22 16:19:31 2020

@author: Farhan
"""


import xlsxwriter
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

def importData():
    arrData = []
    inputan = pd.read_excel ('Mahasiswa.xls')
    for x in range(len(inputan['Id'])):    
        arrData.append([])   
        arrData[x].append(inputan['Id'][x])
        arrData[x].append(inputan['Penghasilan'][x])
        arrData[x].append(inputan['Pengeluaran'][x])
    return arrData

def membershipPenghasilan():
    x1 = [0,3,5,20] 
    y1 = [1,1,0,0] 
    
    x2 = [0,4,5,10,11,20] 
    y2 = [0,0,1,1,0,0]
    
    x3 = [0,10,12,20] 
    y3 = [0,0,1,1]
    
    plt.plot(x1, y1,'r-',label = 'Low Income') 
    plt.plot(x2, y2,'b-',label = 'Mid Income') 
    plt.plot(x3, y3,'g-',label = 'High Income')
      
    plt.title('Penghasilan *satuan dalam juta*') 
    plt.legend()
    plt.xticks(np.arange(min(x1), max(x1)+1, 2.0))
    plt.show()

def membershipPengeluaran():    
    x1 = [0,3,5,20] 
    y1 = [1,1,0,0] 
  
    x2 = [0,4,5,7,8,20] 
    y2 = [0,0,1,1,0,0]
    
    x3 = [0,7,9,20] 
    y3 = [0,0,1,1]
    
    plt.plot(x1, y1,'r-',label = 'Low Outcome') 
    plt.plot(x2, y2,'b-',label = 'Mid Outcome') 
    plt.plot(x3, y3,'g-',label = 'High Outcome') 
    
    plt.title('Pengeluaran *satuan dalam juta*') 
    plt.legend()
    plt.xticks(np.arange(min(x1), max(x1)+1, 2.0))
    plt.show()

def fuzzyPenghasilan(income):
    #low
    if (income <= 3 ) :
        low = 1
    elif (income >= 5) :
        low = 0
    elif (income > 3 and income < 5) :
        low = (5-income)/(5-3)
    #mid
    if (income <= 4 or income >= 11) :
        mid = 0
    elif(income >= 5 and income <= 10) :
        mid = 1
    elif(income > 4 and income < 5):
        mid = (income-4)/(5-4)
    elif(income > 10 and income < 11):
        mid = (11-income)/(11-10)
    #high
    if(income <= 10):
        high = 0
    elif(income >= 12):
        high = 1
    elif(income > 10 and income < 12):
        high = (income-10)/(12-10)
    return round(low,3),round(mid,3),round(high,3)

def fuzzyPengeluaran(outcome):
    #low
    if(outcome <= 3):
        low = 1
    elif(outcome >= 5):
        low = 0
    elif(outcome > 3 and outcome < 5):
        low = (5-outcome)/(5-3)
    #mid
    if (outcome <= 4 or outcome >= 8):
        mid = 0
    elif(outcome >= 5 and outcome <= 7):
        mid = 1
    elif(outcome > 4 and outcome < 5):
        mid = (outcome-4)/(5-4)
    elif(outcome > 7 and outcome < 8):
        mid = (8-outcome)/(8-7)
    #high
    if(outcome <= 7):
        high = 0
    elif(outcome >= 9):
        high = 1
    elif(outcome > 7 and outcome < 9):
        high = (outcome-7)/(9-7)
    return round(low,3),round(mid,3),round(high,3)

def fuzzyRules(penghasilan,pengeluaran):
    arrRules =[
        #low,low
        ['Consider',min(penghasilan[0],pengeluaran[0])],
        #low,mid
        ['Accept',min(penghasilan[0],pengeluaran[1])],
        #low,high
        ['Accept',min(penghasilan[0],pengeluaran[2])],
        #mid,low
        ['Reject',min(penghasilan[1],pengeluaran[0])],
        #mid,mid
        ['Reject',min(penghasilan[1],pengeluaran[1])],
        #mid,high
        ['Accept',min(penghasilan[1],pengeluaran[2])],
        #high,low
        ['Reject',min(penghasilan[2],pengeluaran[0])],
        #high,mid
        ['Reject',min(penghasilan[2],pengeluaran[1])],
        #high,high
        ['Reject',min(penghasilan[2],pengeluaran[2])]]
    
    return arrRules

def inference(arrRules):
    arrCon = []
    arrAcc = []
    arrRej = []
    for x in range(len(arrRules)):
        if(arrRules[x][0] == 'Consider'):
            arrCon.append(arrRules[x][1])
        elif(arrRules[x][0] == 'Accept'):
            arrAcc.append(arrRules[x][1])
        elif(arrRules[x][0] == 'Reject'):
            arrRej.append(arrRules[x][1])
    
    return max(arrAcc),max(arrCon),max(arrRej)

def deFuzzy(arrInference):
    defuzzy = ((arrInference[0]*100)+(arrInference[1]*75)+(arrInference[2]*50))/(arrInference[0]+arrInference[1]+arrInference[2])
    return defuzzy

def getFinalResult(arrResult):
    arrTemp = []
    arrResult = sorted(arrFinalResult, key=lambda x: x[1], reverse=True)
    print()
    for i in range (20):
        arrTemp.append(arrResult[i][0])
    arrTemp.sort(reverse=False)
    return arrTemp

dataMahasiswa = importData()
membershipPenghasilan()
membershipPengeluaran()

arrFuzzHasil = []
for i in range(len(dataMahasiswa)):
    arrFuzzHasil.append(fuzzyPenghasilan(dataMahasiswa[i][1]))
arrFuzzKeluar = []
for i in range(len(dataMahasiswa)):
    arrFuzzKeluar.append(fuzzyPengeluaran(dataMahasiswa[i][2]))

arrFinalResult = []
print()
for i in range(len(dataMahasiswa)):
    fuzzy = fuzzyRules(arrFuzzHasil[i],arrFuzzKeluar[i])
    infer=inference(fuzzy)
    arrFinalResult.append([i+1,deFuzzy(infer)])
arrFinalResult = getFinalResult(arrFinalResult)
print('id:',arrFinalResult)

workbook = xlsxwriter.Workbook('Bantuan.xlsx')
worksheet = workbook.add_worksheet("The Data")
worksheet.write(0,0,'id')

start = 1 
for i in range(20):
    worksheet.write(start,0,arrFinalResult[i])
    start += 1

workbook.close()