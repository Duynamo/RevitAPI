####string method
# name = 'Duy Dep Trai'
# print(len(name))
# print(name.find('D'))
# print(name.capitalize() )
# print(name.upper() )
# print(name.lower())
# print(name.isdigit())
# print(name.isalpha())
# print(name.count('D'))
# print(name.replace('D','C'))
# print(name*3)

####userInput
# userName = input('What is your name? ')
# print('your name is ' + userName + ' , correct? ')

#### math library
# import math
# pi = 3.1444
# print(round(pi))
# print(math.ceil(pi))
# print(math.floor(pi))
# print(pow(pi,5))
# print(math.sqrt(pi))

####slicing 
# name = 'Vu Dinh Duy'
# #[start:stop:step]
# firstName = name[0:2]#==Vu
# firstName1 = name[:2]
# name2 = name[0:11:3]#==VDhu
# funkyName = name[::3]#star == 0 && end == 0
# #reverse string by slicing
# reversedName = name[::-1]

# print(reversedName)

####while loop
#while loop is a statement that will execute it's block of code, as long as it's condition remains true

name = ''
while len(name) == 0:
    name = input('please input your name: ')

print('Hello ' + name)





