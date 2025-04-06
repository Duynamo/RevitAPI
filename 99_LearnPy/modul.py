#NOTE Ta nên gom tất cả các hàm thực thi của file python vào conditional statement

# print(f'__name__ in the modul.py: {__name__}')

def func(a,b):
    print('a+b = ', a+b)

if __name__ == '__main__':
    func(100,1)