####Kiểm tra ước lẻ:
def check_odd_divisor(n):
    if n <= 1:
        return "NO"
    
    for i in range(3, int(n**0.5) + 1, 2):
        if n % i == 0:
            return "YES" if i % 2 == 1 else "NO"
    
    return "NO"

# Đọc số nguyên dương n từ người dùng
n = int(input("Nhập số nguyên dương n: "))

# Kiểm tra và in kết quả
result = check_odd_divisor(n)
print(result)
