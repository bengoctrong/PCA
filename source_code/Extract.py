import xlrd
import numpy as np
from PIL import Image
##################################################################
D = 180*200
K = 100
U1 = np.zeros((D,K))
U2 = np.zeros((D,K))
U3 = np.zeros((D,K))
kv = np.zeros((D,3))
Z = np.zeros((K,3))
##################################################################
input =  r"D:\du lieu\NAM BA\doan_AI\data\train_data.xlsx"
wb  = xlrd.open_workbook(input)
sheet1 = wb.sheet_by_index(0)
sheet2 = wb.sheet_by_index(1)
sheet3 = wb.sheet_by_index(2)
sheet_kv = wb.sheet_by_index(3)
##################################################################
for i in range(D):
    kv[i,0] = sheet_kv.cell_value(i,0)
    kv[i,1] = sheet_kv.cell_value(i,1)
    kv[i,2] = sheet_kv.cell_value(i,2)
    for j in range(K):
        U1[i,j] = sheet1.cell_value(i,j)
        U2[i,j] = sheet2.cell_value(i,j)
        U3[i,j] = sheet3.cell_value(i,j)
##################################################################
path = r'D:\du lieu\NAM BA\doan_AI\data'
wb = xlrd.open_workbook(path+'\\anh.xlsx')
sheet_data = wb.sheet_by_index(0)
##################################################################
for i in range(K):
    for j in range(3):
        Z[i,j] = sheet_data.cell_value(i,j)
##################################################################
red = U1.dot(Z[:,0].reshape(K,1)) + kv[:,0].reshape(D,1)
green = U2.dot(Z[:,1].reshape(K,1)) + kv[:,1].reshape(D,1)
blue = U3.dot(Z[:,2].reshape(K,1)) + kv[:,2].reshape(D,1)       
##################################################################
I = np.zeros((200,180,3))
I[:,:,0] = red.reshape(200,180)
I[:,:,1] = green.reshape(200,180)
I[:,:,2] = blue.reshape(200,180)    
##################################################################
im =  np.zeros((200,180,3),dtype = np.int8) 
for x in range(200):
    for y in range(180):
        for z in range(3):
            im[x,y,z] = int(I[x,y,z])
c = Image.fromarray(im, mode = 'RGB')
c.save(path+"\\anh.jpg")
        
        
        
    