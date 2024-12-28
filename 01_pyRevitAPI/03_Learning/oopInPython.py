#region Class trong OOP là gì?
#Class đơn giản được hiểu là 1 khuôn mẫu, ở đó ta khai báo các thuộc tính (attribute) 
#và các phương thức (method)nhằm miêu tả để từ đó ta tạo ra được những object(đối tượng) mong muốn
#Đôi khi object người ta cũng có thể ghi là instance, tuy nhiên, 
#endregion



#region __OOP without constructor
# class sieuNhan:
#     pass

# sieuNhanA = sieuNhan()
# #sieunhanA được gọi là 1 object và object này có kiểu dữ liệu là sieuNhan
# sieuNhanA.ten = 'Vu Dinh Duy'
# sieuNhanA.vukhi = 'Katana'
# *.vukhi là 1 attribute

# print (sieuNhanA)
# print ('Ten cua sieu nhan A la: ', sieuNhanA.ten)
# print ('Vu khi cua sieu nhan A la: ', sieuNhanA.vukhi)
# print ('Vu khi cua sieu nhan A la: ', sieuNhanA.nhomMau)
# #dòng bên trên khi print sẽ báo lỗi. lý do là không thể dùng 1 attribute chưa từng đưuọc khởi tạo giá trị.
#endregion

#region __OOP with constructor
# class SieuNhan():
#     def __init__(self, para_ten, para_mauSac, para_vuKhi):
#         self.ten = 'Siêu nhân: ' + para_ten
#         self.mauSac = para_mauSac
#         self.vuKhi = para_vuKhi
#     def Hello(self):
#         return 'Xin chào, ta chính là' + self.ten    
#     def Hello1(self):
#         print('Xin chào, ta chính là ' + self.ten)
#     pass
# #endregion
# sieuNhanB = SieuNhan('Duy', 'Blue','Katana')
# #sieuNhanB  là 1 thể hiện của class SieuNhan

# # print (sieuNhanB)
# # print (sieuNhanB.ten)
# # print (sieuNhanB.vuKhi)
# # print (sieuNhanB.Hello())
# # print (SieuNhan.Hello(sieuNhanB))
# sieuNhanB.Hello1() 
#dùng chính thể hiện để chấm đến method/ altribute đã khởi tạo


#region Day2 Thuộc tính của lớp
class SieuNhan:
    sucManh = 10000
    stt = 1
    soThuTu = 1
    def __init__(self, para_ten, para_vuKhi, para_color):
        self.ten = 'Sieu nhan: ' + para_ten
        self.vuKhi = para_vuKhi
        self.color = para_color
        self.stt = SieuNhan.soThuTu + 1
        SieuNhan.soThuTu += 1
        
sieuNhanA = SieuNhan('vitaminD', 'Knowledge', 'Magenta')
sieuNhanB = SieuNhan('vitaminR', 'Run', 'Blue')

print(sieuNhanA.sucManh)
print(SieuNhan.sucManh)

SieuNhan.sucManh = 5

print(SieuNhan.stt)
print(sieuNhanA.stt)
print(sieuNhanB.stt)

#endregion

#region
class SieuNhan:
    sucManh = 100
    def __init__(self, para_ten, para_vuKhi, para_color):
        self.ten = para_ten
        self.vuKhi = para_vuKhi
        self.color = para_color

class SieuNhanVitamin(SieuNhan):
# SieuNhanVitamin đã được kế thừa mọi thuộc tính của SieuNhan
    sucManh = 10000000000000000000000
    def __init__(self, para_ten, para_vuKhi, para_color, para_pet):
        super().__init__(para_ten, para_vuKhi, para_color)
        # hàm super() cho phép SieuNhanVitamin kết thừa mọi thuộc tính của SieuNhan(quá khủng kk)
        self.pet = para_pet


vitaminD = SieuNhanVitamin('Duy', 'Sung', 'RED', 'Lion')
print(vitaminD.__dict__)
#endregion
