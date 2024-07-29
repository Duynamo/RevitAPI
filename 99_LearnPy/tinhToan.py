import math

#region _Bai1  Tinh gia tri cua ham so f(x) = e^x - x tai gia tri x = 0.501
x = 0.501
f = pow(math.e,x) - x
# print (round(f, 2))
#endregion#region _Bai1  Tinh gia tri cua ham so f(x) = e^x - x tai gia tri x = 0.501
# x = float(input('Enter x value:'))
f1 = pow(x, 2)
# print (round(f1, 2))
#endregion
#region _Bai3: Viết chương trình cho phép người dùng nhập vào 1 chuỗi, xuất ra 2 ký tự đầu và cuối của chuỗi. Trong trường hợp chuỗi nhập vào ngắn hơn 2, trả về Empty String
while True:
    inputString = str(input('Hãy nhập chuỗi:'))
    strLength = len(inputString)
    if strLength >=2:
        print (inputString[0:2]+inputString[strLength-2:strLength])
        break
    elif strLength <2 :
        print ('None')

#endregion