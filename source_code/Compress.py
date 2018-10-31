import xlrd
import xlsxwriter
import numpy as np
from PIL import Image
##################################################################
D = 180*200
K = 100
U1 = np.zeros((D,K))
U2 = np.zeros((D,K))
U3 = np.zeros((D,K))
kv = np.zeros((D,3))
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
path = r'D:\du lieu\NAM BA\doan_AI\data\nen.jpg'
img = Image.open(path)
I = np.asarray(img,dtype = np.float)
data1 = I[:,:,0].reshape(D,1)
data2 = I[:,:,1].reshape(D,1)
data3 = I[:,:,2].reshape(D,1)
##################################################################
path_out = r'D:\du lieu\NAM BA\doan_AI\data\anh.xlsx'
workbook = xlsxwriter.Workbook(path_out)
sheet_data = workbook.add_worksheet("Image")
##################################################################
data1 = data1 - kv[:,0].reshape(D,1)
Z = (U1.T).dot(data1)
for i in range(K):
    sheet_data.write_number(i,0,Z[i])
##################################################################
data2 = data2 - kv[:,1].reshape(D,1)
Z = (U2.T).dot(data2)
for i in range(K):
    sheet_data.write_number(i,1,Z[i])
##################################################################
data3 = data3 - kv[:,2].reshape(D,1)
Z = (U3.T).dot(data3)
for i in range(K):
    sheet_data.write_number(i,2,Z[i]) 
##################################################################
workbook.close()