import modul
import argparse

# print(f'__name__ in the run.py: {__name__}')

def my_func(a,b):
    print('a-b = ', a-b)
#NOTE: tại đây, ta run hàm my_func() nhưng giá trị trả về lại bao gồm cả hàm func() của modul.py
#Lý do là tại modul.py ta không đặt hàm func vào một conditonal statement
#Sau khi đặt hàm func trong modul.py vào 1 conditional statement, ta đã thu được kết quả mong muốn
if __name__ == '__main__':
    my_func(100,1)