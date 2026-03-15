# ============================================================
#  PNEUMATICGLOBEVALVE.py  —  Plant 3D 2026
#  Viết theo đúng pattern của SO_Flange
# ============================================================

#region ___import Python Package
from aqa.math import *
from varmain.primitiv import *
from varmain.custom import *
#endregion

#region __metaData
@activate(Group="Valves",
          TooltipShort="Pneumatic Globe Valve",
          TooltipLong="Parametric pneumatic actuated globe valve with flanged ends",
          FirstPortEndtypes="FL",
          LengthUnit="mm",
          Ports="2")
@group("MainDimensions")
@param(D    = LENGTH, TooltipShort="Nominal bore diameter")
@param(L    = LENGTH, TooltipShort="Face-to-face length")
@param(D1   = LENGTH, TooltipShort="Flange outside diameter")
@param(TF   = LENGTH, TooltipShort="Flange thickness")
@param(HB   = LENGTH, TooltipShort="Bonnet height above centreline")
@group("Actuator")
@param(DACT = LENGTH, TooltipShort="Actuator cylinder diameter")
@param(HACT = LENGTH, TooltipShort="Actuator cylinder height")
#endregion

def PNEU_GLOBE_VALVE(s,
                     D    = 100.0,
                     L    = 400.0,
                     D1   = 210.0,
                     TF   =  22.0,
                     HB   = 200.0,
                     DACT = 200.0,
                     HACT = 150.0,
                     **kw):

    # ---- Derived values ----
    R_BORE    = D    / 2.0
    R_FL      = D1   / 2.0
    R_BODY    = D    * 1.15
    R_BONNET  = D    * 0.60
    R_BFLANGE = D    * 0.90
    TBF       = TF   * 0.70
    R_YOKE    = D    * 0.10
    YOKE_H    = HB   * 0.35
    YOFF      = DACT * 0.38
    R_ACT     = DACT / 2.0
    H_ACT_BDY = HACT * 0.78
    H_ACT_CAP = HACT * 0.12
    H_ACT_AFL = HACT * 0.10
    R_STEM    = D    * 0.11
    H_STEM    = HACT * 0.22
    CX        = L    / 2.0

    # ---- 1. Body + flanges ----
    o_body = CYLINDER(s, R=R_BODY, H=L - 2.0*TF, O=R_BORE).rotateY(90).translate((TF, 0, 0))
    o_fl1  = CYLINDER(s, R=R_FL,   H=TF,          O=R_BORE).rotateY(90)
    o_fl2  = CYLINDER(s, R=R_FL,   H=TF,          O=R_BORE).rotateY(90).translate((L - TF, 0, 0))
    o_body.uniteWith(o_fl1); o_fl1.erase()
    o_body.uniteWith(o_fl2); o_fl2.erase()

    # ---- 2. Bonnet ----
    o_bonnet = CYLINDER(s, R=R_BONNET,  H=HB,  O=0).translate((CX, 0, 0))
    o_bfl    = CYLINDER(s, R=R_BFLANGE, H=TBF, O=0).translate((CX, 0, HB - TBF))
    o_bonnet.uniteWith(o_bfl);   o_bfl.erase()
    o_body.uniteWith(o_bonnet);  o_bonnet.erase()

    # ---- 3. Yoke ----
    CROSS_L = 2.0 * YOFF + 2.0 * R_YOKE
    o_yk1   = CYLINDER(s, R=R_YOKE, H=YOKE_H, O=0).translate((CX - YOFF, 0, HB))
    o_yk2   = CYLINDER(s, R=R_YOKE, H=YOKE_H, O=0).translate((CX + YOFF, 0, HB))
    o_cross = CYLINDER(s, R=R_YOKE, H=CROSS_L, O=0).rotateY(90).translate(
                       (CX - YOFF - R_YOKE, 0, HB + YOKE_H - R_YOKE))
    o_body.uniteWith(o_yk1);   o_yk1.erase()
    o_body.uniteWith(o_yk2);   o_yk2.erase()
    o_body.uniteWith(o_cross); o_cross.erase()

    # ---- 4. Pneumatic actuator ----
    TOP_YOKE = HB + YOKE_H
    o_afl = CYLINDER(s, R=R_ACT * 1.10, H=H_ACT_AFL, O=0).translate((CX, 0, TOP_YOKE - H_ACT_AFL * 0.5))
    o_act = CYLINDER(s, R=R_ACT,         H=H_ACT_BDY, O=0).translate((CX, 0, TOP_YOKE))
    o_cap = CYLINDER(s, R=R_ACT * 1.04,  H=H_ACT_CAP, O=0).translate((CX, 0, TOP_YOKE + H_ACT_BDY))
    o_air = CYLINDER(s, R=D * 0.07, H=R_ACT * 0.55, O=0).rotateX(-90).translate(
                     (CX, -R_ACT, TOP_YOKE + H_ACT_BDY * 0.40))
    o_body.uniteWith(o_afl); o_afl.erase()
    o_body.uniteWith(o_act); o_act.erase()
    o_body.uniteWith(o_cap); o_cap.erase()
    o_body.uniteWith(o_air); o_air.erase()

    # ---- 5. Stem ----
    o_stem = CYLINDER(s, R=R_STEM, H=H_STEM, O=0).translate(
                      (CX, 0, TOP_YOKE + H_ACT_BDY + H_ACT_CAP))
    o_body.uniteWith(o_stem); o_stem.erase()

    return None

# ---- Dòng bắt buộc để Plant 3D đăng ký script ----
valve = PNEU_GLOBE_VALVE
OUT   = valve
# ```

# ---

# ## Sau khi lưu file
# ```
# PLANTREGISTERCUSTOMSCRIPTS
# ```

# Kiểm tra lại `variants.map` — phải thấy dòng:
# ```
# PNEU_GLOBE_VALVE;customscripts.pneumaticglobevalve