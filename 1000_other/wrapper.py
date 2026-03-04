from varmain.custom import * # type: ignore
import importlib  # ✅ FIX: Cần import để dùng reload trong Python 3

# Import primitives for fallback geometry
from primitives import Box

# import the real script module
import test_geometry

#(testacpscript "wrapper")
@activate() # type: ignore
def wrapper(s, D=80, **kw):
    try:
        # ✅ FIX: Python 3 dùng importlib.reload(), không dùng reload() trực tiếp
        importlib.reload(test_geometry)
        
        result = test_geometry.build_geometry_test(s)
        return result

    except Exception as e:
        print("ERROR in test_geometry: " + str(e))
        fallback_geom = Box(s, L=D, W=D, H=D)
        return fallback_geom