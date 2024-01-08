class SieuNhan:
    suc_manh_tong = 1000
    def __init__(self, para_ten, para_vuKhi, para_sucManh):
        self.ten = "Sieu Nhan " + para_ten
        self.vuKhi = para_vuKhi
        self.sucManh = para_sucManh
    def xin_chao(self):
        return "Xin chao " + self.ten
     
sieu_nhan_A = SieuNhan("Duy","Sung",8000)

print(sieu_nhan_A.xin_chao())
# print(sieu_nhan_A.ten(), sieu_nhan_A.vuKhi())
print("Vu khi cua sieu nhan A la " + sieu_nhan_A.vuKhi)
print(sieu_nhan_A.suc_manh_tong)
print(SieuNhan.suc_manh_tong)

SieuNhan.suc_manh_tong = 9000000
sieu_nhan_A.sucManh = 80000000000000
print(sieu_nhan_A.sucManh)
print(SieuNhan.suc_manh_tong)