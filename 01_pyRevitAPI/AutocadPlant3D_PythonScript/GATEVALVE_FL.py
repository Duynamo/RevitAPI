# ============================================================
#  GATEVALVE_FL.py  —  Plant 3D 2026
#  Flanged Gate Valve — 3 nan hoa cách đều 120°
# ============================================================

#region ___import Python Package
from aqa.math import *
from varmain.primitiv import *
from varmain.custom import *
#endregion

#region __metaData
@activate(Group="Valves",
          TooltipShort="Gate Valve Flanged",
          TooltipLong="Parametric flanged gate valve with 3-spoke handwheel",
          FirstPortEndtypes="FL",
          LengthUnit="mm",
          Ports="2")
@group("MainDimensions")
@param(D   = LENGTH, TooltipShort="Nominal bore diameter")
@param(L   = LENGTH, TooltipShort="Face-to-face length")
@param(D1  = LENGTH, TooltipShort="Flange outside diameter")
@param(TF  = LENGTH, TooltipShort="Flange thickness")
@param(HB  = LENGTH, TooltipShort="Bonnet height above centreline")
@group("Handwheel")
@param(DHW = LENGTH, TooltipShort="Handwheel diameter")
@param(THW = LENGTH, TooltipShort="Handwheel thickness")
#endregion

def GATE_VALVE_FL(s,
                  D   = 100.0,
                  L   = 360.0,
                  D1  = 210.0,
                  TF  =  22.0,
                  HB  = 220.0,
                  DHW = 200.0,
                  THW =  14.0,
                  **kw):

    # ---- Derived values ----
    R_BORE    = D   / 2.0
    R_FL      = D1  / 2.0
    R_BODY    = D   * 1.20
    R_BONNET  = D   * 0.55
    R_BFLANGE = D   * 0.85
    TBF       = TF  * 0.65
    R_STEM    = D   * 0.09
    H_STEM    = HB  * 0.25
    R_HW      = DHW / 2.0
    R_HUB     = R_STEM * 2.5
    H_HUB     = THW * 1.8
    CX        = L   / 2.0

    # ---- 1. Body + 2 mặt bích ----
    o_body = CYLINDER(s, R=R_BODY, H=L - 2.0*TF, O=R_BORE
                      ).rotateY(90).translate((TF, 0, 0))
    o_fl1  = CYLINDER(s, R=R_FL, H=TF, O=R_BORE).rotateY(90)
    o_fl2  = CYLINDER(s, R=R_FL, H=TF, O=R_BORE
                      ).rotateY(90).translate((L - TF, 0, 0))
    o_body.uniteWith(o_fl1); o_fl1.erase()
    o_body.uniteWith(o_fl2); o_fl2.erase()

    # ---- 2. Bonnet flange ----
    o_bfl = CYLINDER(s, R=R_BFLANGE, H=TBF, O=0
                     ).translate((CX, 0, R_BODY - TBF * 0.3))
    o_body.uniteWith(o_bfl); o_bfl.erase()

    # ---- 3. Bonnet neck ----
    o_bonnet = CYLINDER(s, R=R_BONNET, H=HB - R_BODY, O=0
                        ).translate((CX, 0, R_BODY))
    o_body.uniteWith(o_bonnet); o_bonnet.erase()

    # ---- 4. Stem ----
    STEM_BASE_Z = HB
    o_stem = CYLINDER(s, R=R_STEM, H=H_STEM, O=0
                      ).translate((CX, 0, STEM_BASE_Z))
    o_body.uniteWith(o_stem); o_stem.erase()

    # ---- 5. Handwheel hub ----
    HW_Z  = STEM_BASE_Z + H_STEM - H_HUB / 2.0
    o_hub = CYLINDER(s, R=R_HUB, H=H_HUB, O=0
                     ).translate((CX, 0, HW_Z))
    o_body.uniteWith(o_hub); o_hub.erase()

    # ---- 5b. Spokes — 3 nan hoa cách đều 120° ----
    R_SPOKE = R_STEM * 0.80
    SPOKE_Z = HW_Z + (H_HUB - R_SPOKE * 2.0) / 2.0

    o_sp1 = CYLINDER(s, R=R_SPOKE, H=2.0*R_HW, O=0
                     ).rotateY(90
                     ).translate((-R_HW, 0, 0)
                     ).rotateZ(0
                     ).translate((CX, 0, SPOKE_Z))

    o_sp2 = CYLINDER(s, R=R_SPOKE, H=2.0*R_HW, O=0
                     ).rotateY(90
                     ).translate((-R_HW, 0, 0)
                     ).rotateZ(120
                     ).translate((CX, 0, SPOKE_Z))

    o_sp3 = CYLINDER(s, R=R_SPOKE, H=2.0*R_HW, O=0
                     ).rotateY(90
                     ).translate((-R_HW, 0, 0)
                     ).rotateZ(240
                     ).translate((CX, 0, SPOKE_Z))

    o_body.uniteWith(o_sp1); o_sp1.erase()
    o_body.uniteWith(o_sp2); o_sp2.erase()
    o_body.uniteWith(o_sp3); o_sp3.erase()

    # ---- 6. Handwheel rim ----
    o_hw = CYLINDER(s, R=R_HW, H=THW, O=R_HW * 0.70
                    ).translate((CX, 0, HW_Z + (H_HUB - THW) / 2.0))
    o_body.uniteWith(o_hw); o_hw.erase()

    # ---- Connection points ----
    s.setPoint((0, 0, 0), (-1, 0, 0))
    s.setPoint((L, 0, 0), ( 1, 0, 0))

    return None

# ---- Bắt buộc để Plant 3D đăng ký ----
valve = GATE_VALVE_FL
OUT   = valve
