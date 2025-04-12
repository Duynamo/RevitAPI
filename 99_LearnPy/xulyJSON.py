import json

dict = {
    'name' : 'Duy',
    'age': 1000,
    'loc' : 'Da Nang',
    'hobbies' : ['eating', 'playing', 'reading']
}

json_string = json.dumps(dict) #lưu ý, khi convert từ dict sang json , phương thức dumps() thì s đại diện cho string
# print(type(json_string), json_string)

new_dict  = json.loads(json_string)
# print(type(new_dict), new_dict)

# assert new_dict == dict #Kiểm tra xem dict ban đầu và sau khi chuyển qua lại có giống nhau hay không
# print('2 dict nay giong nhau')

#NOTE: ghi python vào file .json
# with open ('save.json', 'w') as f:
#     json.dump(dict, f, indent = 5) #indent là thụt dùng cho dễ nhìn

#NOTE: đọc .json file về python object
with open('save.json') as f: #ở đây tự hiểu là đọc file nên ko cần thêm 'w'
    my_content = json.load(f)

print(type(my_content), my_content)