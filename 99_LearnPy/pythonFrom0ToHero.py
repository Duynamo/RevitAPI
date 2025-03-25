#region String

#index
myString = 'Welcome To VietNam'
a = myString [0] 
b = myString [-1] #-> Lấy về giá trị thứ  len(myString) - 1
c = myString [17]

#slicing
sl1 = myString[0:7] # lấy về chuỗi có index từ 0 -> i-1
sl2 = myString[:-1]
#string immutable
# myString[0] = 'D' -> không thể thay thế ký tự chuỗi
myString2 = ['DUY']
# myString2[0] = 'W'


#endregion

#region Toán tử
x = True
y = False

#endregion

#region built-in data type _list, tuple, set , dictionary
myList = [1,2,1,100,-1000]
myList.sort(reverse=True)
newList = sorted(myList)
newList2 = myList[::-1] #start:stop:step

#insert element to list
myList.insert(0, "sss")

#delete of remove item in list
del myList[0]
del myList[0:2]
newList5 = [5,66,99,65540]
newList5.remove(66)
print (newList5)

#index
newList3 = [1,1,8,19,64]
print (newList3.index(1)) # trả về index của phần tử đầu tiên được khớp .index()

#pop()
newList4 = [4,8,66]
popEle = newList4.pop(0)

print (newList4, popEle)
#endregion