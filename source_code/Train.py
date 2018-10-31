import xlsxwriter
import math
from PIL import Image
import numpy as np
from numpy.linalg import eigh
##################################################################
def norm2_equ1(x):
    for i in range(x.shape[0]):
        k = 0
        for j in range(x.shape[1]):
            k += x[i,j]*x[i,j]
        for j in range(x.shape[1]):
            x[i,j] /= math.sqrt(k)
    return x
##################################################################
D = 200*180
N = 200  
K = 100
X1 = np.zeros((D,N))
X2 = np.zeros((D,N))
X3 = np.zeros((D,N))
input =  r'D:\du lieu\NAM BA\doan_AI\faces95'
##################################################################
output = r'D:\du lieu\NAM BA\doan_AI\data\train_data.xlsx'
wb = xlsxwriter.Workbook(output)
sheet1 = wb.add_worksheet('U1')
sheet2 = wb.add_worksheet('U2')
sheet3 = wb.add_worksheet('U3')
sheet_kv = wb.add_worksheet('Ky vong')
##################################################################
def Buoc23(x):
    kv = np.mean(x,axis = 1) #tinh vector ky vong
    _x = x - kv.reshape(D,1) #tinh ma tran X^
    T = _x.T.dot(_x) #tinh ma tran T
    T /= 200
    return kv,_x,T
def Buoc456(T):
    evl,evt = eigh(T) 
    #evl la vector chua cac tri rieng, evt la ma tran vector rieng
    evt = evt.T
        
    li = list(zip(evl,evt))
    li.sort(reverse=True) #sort theo gia tri cao -> thap
    
    _evt = np.zeros((K,D))
    
    for i in range(K):
        _evt[i] = _X.dot(li[i][1]) #nhan voi X de duoc vector rieng ung voi S
    _evt = norm2_equ1(_evt) #chuyen thanh cac vector co norm 2 = 1
    U = _evt.T
    return U
##################################################################
#Buoc 1
#Doc cac file anh
index = 0
for i in range(20):
    for j in range(10):
        img = Image.open(input+'\\'+str(i+1)+'\\'+str(i+1)+' ('+str(j+1)+').jpg')
        tmp = np.asarray(img,dtype = np.float64)
        X1[:,index] = tmp[:,:,0].reshape(D)
        X2[:,index] = tmp[:,:,1].reshape(D)
        X3[:,index] = tmp[:,:,2].reshape(D)
        index += 1

kv,_X,T = Buoc23(X1)
U1 = Buoc456(T)
for i in range(D):
    sheet_kv.write_number(i,0,kv[i])
    for j in range(K):
        sheet1.write_number(i,j,U1[i,j])
##################################################################
kv,_X,T = Buoc23(X2)
U2 = Buoc456(T)
for i in range(D):
    sheet_kv.write_number(i,1,kv[i])
    for j in range(K):
        sheet2.write_number(i,j,U2[i,j])
##################################################################
kv,_X,T = Buoc23(X3)
U3 = Buoc456(T)
for i in range(D):
    sheet_kv.write_number(i,2,kv[i])
    for j in range(K):
        sheet3.write_number(i,j,U3[i,j])
wb.close()