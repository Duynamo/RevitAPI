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

#region cau truc if else
# num1 = int(input('Nhap so can kiem tra: '))
# if num1 % 2 == 0:
#     print(f'so can kiem tra {num1} la so chan')
# else:
#     print (f'so can kiem tra {num1} la so le')
#endregion

#region loop
#for loop -> dùng khi biết trước số lần lặp
# list1 = [1,2,3,4,5,6]
# for i in range(len(list1)):
#     print(f'{i} : {list1[i]}')

#while loop -> dùng khi chưa biết trước số lần lặp, và lặp đến khi điều kiện còn đúng
# checkNum = 1
# while checkNum < 5 :
#     print ('ok')
#     checkNum += 1

#break and continue. 
list2 = [1,2,3,4]
# for i in range(len(list2)):
    # if i == 5:
    #     break
    # else:
    #     print(i)
    # if i == 3 :
    #     continue
    # print(f'{i} : {list2[i]}')

#endregion

#region def
import shutil
import os

def copyFile(sourceDir, desDir):
    list_name = os.listdir(sourceDir)

    for file_name in list_name:
        shutil.copy(os.path.join(sourceDir, file_name), os.path.join(desDir, file_name))
    return 
# des_dir = r"C:\Users\Laptop\OneDrive\Desktop\testFolder2"
# source_dir = r"C:\Users\Laptop\OneDrive\Desktop\testFolder1"
# a = copyFile(source_dir , des_dir)

#endregion
#region _file handling in python
#region _ đọc file
#thủ công
# file_obj1 = open(r'C:\Users\Laptop\OneDrive\Desktop\testFolder2\testABC.txt')
# content = file_obj1.read()
# print(content)
# file_obj1.close()

#context manager
# with open(r"C:\Users\Laptop\OneDrive\Desktop\testFolder2\testABC.txt") as file_object2:
    content = file_object2.read()
    # print (content)
#endregion

#region _ghi file
# import os

# # Đường dẫn thư mục và tên tệp
# folder_path = r"C:\Users\Laptop\OneDrive\Desktop\testFolder2"
# file_name = "newfile.txt"
# full_path = os.path.join(folder_path, file_name)

# # Đảm bảo thư mục tồn tại
# if not os.path.exists(folder_path):
#     os.makedirs(folder_path)
#     print(f"Created directory: {folder_path}")

# # Ghi vào tệp
# with open(full_path, 'w') as f:
#     f.write("Hello world")
#     print(f"Wrote 'Hello world' to {full_path}")

# #endregion

# #region _appen vào file
# folder_path1 = r"C:\Users\Laptop\OneDrive\Desktop\testFolder2"
# file_name1  =  'testABC - Copy.txt'
# full_path1 = os.path.join(folder_path1,file_name1)
# with open(full_path1 , 'a') as f:
#     f.write('\n Duy dep trai')
#endregion

#region List comprehension
#1 Tổng quan
#print ([x**x for x in [1,2,3,4]])

#2 Filter element
#[expression for item in iterable if condition]
# print([x**2 for x in [1,2,23,4444,22,22222,45465] if x%2 ==0])

#3 apply function to each element
#[expresion1 if condition == True else expresion2 for item in iterable]
# print ([x + 1 if x % 2 == 0 else x + 2 for x in [1,2,3,4,5]])

#4 Dictionary comprehension
#[k:v for item in iterable]
# print ({k: k+10 for k in [1,2,3,4,5,6]})

#5 Set comprehension
#{item for item in iterable}
# print({x for x in 'Duy dep trai'})

def flatten(listA):
    """
    Flatten all nested lists or tuples within listA into a single flat list.
    """
    if not isinstance(listA, (list, tuple)):
        raise ValueError("Input must be a list or tuple")
    
    flat_list = []
    for item in listA:
        if isinstance(item, (list, tuple)):
            flat_list.extend(flatten(item))
        else:
            flat_list.append(item)
    return flat_list

flat1 = 1
# Ví dụ
# print(flatten([1, (2, [3, 4]), 5])) 

#endregion

#endregion
#region xu ly ngoai le
# try:
#     num = int(input('Nhap so A: '))
#     result = 100/num
# except ZeroDivisionError:
#     print('ko the thuc hien phep chia cho 0')
#     print(int(input('Nhap lai so A: ')))
# except ValueError:
#     print('so vua nhap ko hop le')
#     print(int(input('Nhap lai so A: ')))
# else:
#     print(f'so vua nhap hop le, ket qua phep chia la{result}')
# finally:
#     print('xin chao va hen gap lai')
#endregion
#region OOP
class person:
    #class attribute - đặc trưng toàn bộ class
    count = 0
    #hàm khởi tạo
    def __init__(self, name, age):
        self.name = name #self.name chính là instance attribute gắn liền với từng đối tượng
        self.age = age
        person.count += 1
    #định nghĩa method
    def printOut(self):
        print(f'Name : {self.name} Age:{self.age}') 

#Tạo các đối tượng
per1 = person('Duy', 100)
per2 = person('Tuc', 1000)        

# print(person.count)

# print(f'Thong tin nguoi thu nhat la: {per1.name}, tuoi la {per1.age}')
#endreion


#endregion

#region Tính chất đóng gói
# class student:
#     def __init__(self, name, age):
#         self.age = age
#         self.name = name

# myStu1 = student('An', 600)
# print('Truoc thay doi', myStu1.name)
# myStu1.name = 'Duy'
# print('Sau thay doi 1', myStu1.name)

# khi ta thêm _ trước tên biến, ko có gì thay đổi
# class student:
#     def __init__(self, name, age):
#         self.age = age
#         self._name = name

# myStu1 = student('An', 600)
# print('Truoc thay doi', myStu1._name)
# myStu1._name = 'Long'
# print('Sau thay doi 1', myStu1._name)

# class student:
#     def __init__(self, name, age):
#         self.age = age
#         self.__name = name

# myStu1 = student('An', 600)
# # print('Truoc thay doi', myStu1.__name) #Báo lỗi. 

# print('Truoc thay doi', myStu1._student__name) #Đây là quy tắc Name magling syntax đúng sẽ là Tên biến._Class__parameter

#endregion

#region inheritane
class Person:
    def __init__(self, name):
        self.name = name
    def printName(self):
        print(self.name)
class Student(Person): #class con
    # def __init__(self, name, age):
    #     self.name = name
    #     self.age = age
    #tái sử dụng hàm constructor của class cha:
    def __init__(self, name, age):   
        # super().__init__(name) #cách 1 dùng super()
        Person.__init__(self, name) #cách 2 dùng tên class

Student1 = Student('Duy', 100)
# Student1.printName() #ở đây ta thấy, class Student đã kế thừa method printName của class cha (class Person)


#endregion

#region Tính Đa hình _POLYMORPHISM

# class Cow:
#     def __init__(self, name):
#         self.name = name

#     def tiengKeu(self):
#         return print(self.name + ' keu umboooooo')

# class Cat:
#     def __init__(self, name):
#         self.name = name
    
#     def tiengKeu(self):
#         return print(self.name + ' keu meomeo')
    
# conBo = Cow('RedBull')
# conMeo = Cat ('Tom')

# conBo.tiengKeu()
# conMeo.tiengKeu()

# class Animal():
#     def giongSua(self):
#         return 'some sound'
# class Dog(Animal): #ở đây class Dog đã kế thừa và overwritting lại thuộc tính của class cha Animal
#     def giongSua(self):
#         return 'gau gau'

# myAnimal = Animal()
# myDog = Dog()

# print(myAnimal.giongSua())
# print(myDog.giongSua())

#endregion

#region Tính trừu tượng _ Abstraction
from abc import ABC, abstractmethod
#lớp trừu tượng
class Animal (ABC):
    #phương thức trừu tượng
    @abstractmethod
    def speak(self):
        pass
    #ko phải phương thức trừu tượng
    def not_abstract_method(self):
        return 'Khong can ghi de lai'

# ani = Animal() # sẽ báo lỗi vì không thể khởi tạo trực tiếp đối tượng từ class trừu tượng

class Dog(Animal):
    def speak(self):
        return 'Gau Gau'
    
myDog = Dog()
# print(myDog.speak()) #Sau khi khởi tạo lớp kế con kế thừa từ lớp cha trừu tượng , ta đã có thể thao tác
# print(myDog.not_abstract_method())
#endregion

#region _Higher order function
#1. Nhận functions như tham số

# def cal_sum(nums):
#     return sum(nums)
# def higher_order_func(my_func, my_list):
#     my_sum = my_func(my_list)
#     return my_sum

# result1 = higher_order_func(cal_sum, [1,5,9])

# input_str = input("Hay nhap chuoi so (cach nhau bang dau ,): ")  # Ví dụ: "1 2 3"
# inputList = [int(x) for x in input_str.split(',')]
# result2 = higher_order_func(cal_sum, inputList)
# print(f'Tong cua chuoi da nhap la: {result2}')
#2. Trả về function (như kết quả trả về)
#Trả về hàm trong hàm
def cal_min(a,b):
    if a<b:
        return a
    else:
        return b
def cal_max(a,b):
    if a<b :
        return b
    else:
        return a
#Tạo higher order function và trong này sẽ trả về hàm thông thường
def higher_order_function(cal_name):
    if cal_name == 'min':
        return cal_min      #trả về tên của function đã định nghĩa ở trên cal_min
    elif cal_name == 'max':
        return cal_max      #trả về tên của function đã định nghĩa ở trên cal_max
    
#Note: muốn đi tính min của 2 số
myRes1 = higher_order_function('min')
myRes2 = higher_order_function('max')
# print(myRes1(1,1000))
# print(myRes2(1,1000))


#endregion

#region _ Decorator
#1. dùng higher-order function thông thường
#wrapper (hàm bao) và wrapped (hàm được bao)

#Dùng higer-order function để thay đổi tính năng của wrapped
# def decorator_example(func):
#     def wrapper():
#         print('Before the function is called')
#         func()
#         print('After the function is called')

#     return wrapper
#@decorator_example # chỉ sử dụng khi áp dụng decorator
# def wrapped(): #hàm được bao
#     print ('hello')

# myFunc = decorator_example(wrapped)
# myFunc()

#NOTE: sử dụng decorator
def decorator_example(func):
    def wrapper():
        print('Before the function is called')
        func()
        print('After the function is called')

    return wrapper
@decorator_example
def sayHello(): #hàm được bao
    print ('hello')

'''
sayHello nó sẽ tương đương với decorator_example(sayHello)
sayHello = decorator_example(sayHello)
'''

#NOTE: sử dụng higher-order func nguyên bản
# myFunc = decorator_example(sayHello)
# myFunc()

#NOTE: sử dụng decorator
# sayHello() #Gọi trực tiếp hàm wrapped (hàm được bao)


#Luyện tập Decorator có đối số
def decorator_func(func):
    #hàm bao
    def wrapper(name, age):
        print('Before the fuction is called')
        func(name, age)
        print('After the functino is called')
    return wrapper

#Hàm được bao
@decorator_func
def say_hello(name, age):
    print(f'Hello {name}, you are {age} years old.')

# myFunc = decorator_func(say_hello)
# myFunc('Duy', 600)

# say_hello('Long', 65656)
#endregion

#region _Format string
#NOTE: String modulo operator %
name = 'Duy'
weigh = 9.9
# my_str = 'my name is %s. I am %s kg' %(name, weigh)
my_str = 'my name is %s. I am %.2f kg' %(name, weigh)
# print(my_str)
#NOTE: .format()
name1 = 'Diu'
location1 = 'danang'
# print('My name is {}. I am living in {}.'.format(name1, location1)) #.format đưuọc truyền vào 1 tuple, có index
# print('My name is {1}. I am living in {0}.'.format(name1, location1)) # Đảo giá trị index được
# print('My name is {x}. I am living in {y}.'.format(x = name1, y = location1)) 

dict1 = {'name' : 'Duy', 'location1': 'Danang'}
# print('My name is {myDict[name]}, I am from {myDict[location1]}'.format(myDict =dict1))

acc = 0.22652654
# print('accuracy: {:.3f}'.format(acc)) #lấy 3 chữ số thập phân sau dấu phẩy

#NOTE: f string .  cách này thường được sử dụng nhất hiện nay
name2 = 'LONG'
loc2 = 'HaNoi'
acc2  = 0.369845875454894564654654856 
print(f'My name is {name2}. I am living in {loc2}!. Rat vui duoc gap ban')
print(f'Do chinh xac la {acc:.5f}!')
#NOTE: Template string
from string import Template
name3 = 'Linh'
t = Template('Hello, $name3.')
print(t.substitute(name3 = name3))
#endregion