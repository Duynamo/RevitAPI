
#region ___import Python Package
from aqa.math import * #type: ignore
from varmain.primitiv import * #type: ignore
from varmain.custom import * #type: ignore
#endregion

#region __metaData
@activate(Group="Valves",
          TooltipShort="Gate Valve Flange",
          TooltipLong="Parametric flanged gate valve",
          FirstPortEndtypes="FL",
          LengthUnit="mm",
          Ports="2") #type: ignore

@group("MainDimensions") #type: ignore
@param(D   = LENGTH, TooltipShort="Valve Norminal Diameter")  #type: ignore
@param(D1  = LENGTH, TooltipShort="Flange outside diameter") #type: ignore
@param(L   = LENGTH, TooltipShort="Face-to-face length") #type: ignore
@param(TF  = LENGTH, TooltipShort="Flange thickness") #type: ignore
@param(H  = LENGTH, TooltipShort="Bonnet height above centreline") #type: ignore
@group("Handwheel") #type: ignore
@param(DHW = LENGTH, TooltipShort="Handwheel diameter") #type: ignore
@param(THW = LENGTH, TooltipShort="Handwheel thickness") #type: ignore

def GATEVALVE(s,
                  D   = 100.0,
                  L   = 360.0,
                  D1  = 210.0,
                  TF  =  22.0,
                  H  = 220.0,
                  DHW = 200.0,
                  THW =  14.0,
                  **kw):
    
    # ---- Derived values ----
    # Thân valve
    D_Body = D * 1.2 #Đường kính thân valve
    H_1 = D1 / 2 # Chiều cao từ tâm đế mặt dưới nắp thân valve
    H_2 = TF * 3 #Chiều dày 2 mặt bích thân vavle
    H_3 = H - (H_1 + H_2 + D_Body / 2) # Chiều cao tay vặn
    L_Body = 3*L / 5 #Chiều dày Body vavlve
    L_BodyFlange = D1 # Chiều dài nắp valve

    #Cụm đỡ trục valve
    Support1_L = L_Body*1.2
    Support1_W = Support1_L / 3
    Support1_H = 20
    Support2_L = L_Body
    Support2_H = 20    
    Support3_