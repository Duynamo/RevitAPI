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
# print (newList5)

#index
newList3 = [1,1,8,19,64]
# print (newList3.index(1)) # trả về index của phần tử đầu tiên được khớp .index()

#pop()
newList4 = [4,8,66]
popEle = newList4.pop(0)

# print (newList4, popEle)

#region tuple
tupleA = ('a', 1, [111, 55])
tupleB = ('b') #nếu khởi tuple 1 phần tử mà được kí hiệu như bên trái, sẽ không thu được tuple mà chỉ thu được string
tupleC = ('c', )

#trong python thì tuple ít được hỗ trợ hơn list, ta có thể chuyển tuple thành list , thao tác, và chuyển ngược lại
tupleD = (1,2,3,4)
list_tupleD = list(tupleD)
list_tupleD.append('Vu Dinh Duy')
list_tupleD.remove(2)
new_tupleD = tuple(list_tupleD)
#hai tuple có cùng giá trị thì có cùng id, còn 2 list có cùng giá trị thì khác id 
tupleAA = (1,2,3,4)
tupleBB = (1,2,3,4)
listAA = [1,2,3,4]
listBB = [1,2,3,4]
#không thể thay đổi giá trị của 1 phần tử trong tuple, nhưng có thể can thiệp vào phần tử con của nó nếu nó là list
tupleDD = ([1,2,3], 'AA', 'BB')
tupleDD[0][0] = 'Duy Dep Trai'
#endregion

#region Dictionary
#khởi tạo Dic
dic1 = {'Duy': 10 , 'Tai' : 8}
newDic = {}
newDic1 = dict
dic1['Duy'] = 100
dic1['Nhan'] = 6 #thêm cặp key-val vào cuối dic
#duyệt qua các phần tử trong dic
dic2 = {'Duy': 10 , 'Tai' : 8, 'Long':555}
# for k in dic2:
#     print(k)
# for k in dic2.keys():
#     print(k)   
# for k in dic2.values():
#     print(k)     
# for k,v in dic2.items():
#     print (k,v)
#endregion

#region Set
#00 tạo Set
mySet = {1, 2, 's'}
emptySet1 = set()
emptySet0 = {} # cách khởi tạo này sẽ sinh ra dictionary rỗng chứ ko sinh ra Set
#01 Không cho phép các phần tử trùng lặp
mySet2 = {'a', 'a', 'b', 'b','c'}
# -> print (mySet2) == {'a', 'b', 'c'}
#02 Check Set là unoder, tức là kết quả sau mỗi lần duyện qua set sẽ thu được vị trí phân tử khác nhau


#03 update hay thêm phần tử vào Set
mySet3= {'a', 'a', 'b', 'b','c', 's'}
mySet3.add ('SSSSSS')
mySet3.discard('a') # loại phần tử 'a' khỏi Set. có thể dùng remove() nhưng
#nếu phần tử ta đưa vào ko có trong Set sẽ ném về 1 Exception, ta phải xử lý thêm

#04 Thao tác với 2 hay nhiều set, biểu đồ ven
my_Set_1 = {1,2,3,4,5,6}
my_Set_2 = {1,2,35,9,7,6}
intersectionSet = my_Set_1.intersection(my_Set_2) #lấy giao 2 Set
unionSet = my_Set_1.union(my_Set_2) #Lấy hợp 2 Set
symmetricSet = my_Set_1.symmetric_difference(my_Set_2) #lấy các phần tử chỉ có ở riêng mỗi set
#endregion
print(symmetricSet)
#endregion