import math

def calculateAN(AD_val, discount, ajio):
    aeValue = round((AD_val * discount) / 100)
    afValue = AD_val - aeValue
    agValueNum = 0.18 if afValue > 2499 else 0.05
    ahValue = afValue - (afValue / (1 + agValueNum))
    ajValue = round((afValue * ajio) / 100)
    
    akValue = AD_val - aeValue - ahValue - ajValue
    alValueNum = 0.18 if akValue > 2499 else 0.05
    amValue = akValue * alValueNum
    return akValue + amValue

def solveForAD_High(targetAN, discount, ajio):
    low = 0
    high = targetAN * 20
    for _ in range(100):
        mid = (low + high) / 2
        if calculateAN(mid, discount, ajio) <= targetAN:
            low = mid
        else:
            high = mid
    return low

target = 756
discount = 65
ajio = 34

res = solveForAD_High(target, discount, ajio)
print(f'Raw Res: {res}')
print(f'Rounded Res: {round(res)}')
print(f'Settlement for {round(res)}: {calculateAN(round(res), discount, ajio)}')
