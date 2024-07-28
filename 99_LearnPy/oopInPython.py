#region __OOP without constructor
# class sieuNhan:

#     pass

# sieuNhanA = sieuNhan()
# #sieunhanA được gọi là 1 object và object này có kiểu dữ liệu là sieuNhan
# sieuNhanA.ten = 'Vu Dinh Duy'
# sieuNhanA.vukhi = 'Katana'

# print (sieuNhanA)
# print ('Ten cua sieu nhan A la: ', sieuNhanA.ten)
# print ('Vu khi cua sieu nhan A la: ', sieuNhanA.vukhi)
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

k = 100
lst = [100,1,2,100]
lst2 = []

lst2 =[c for c in lst if c not in lst2]
print (lst2)