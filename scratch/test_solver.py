import math

def calculateAN(AD_val, discount, ajio):
    aeValue = round((AD_val * discount) / 100)
    afValue = AD_val - aeValue
    agValueNum = 0.18 if afValue > 2499 else 0.05
    ahValue = afValue - (afValue / (1 + agValueNum))
    ajValue = round(((afValue * ajio) / 100) * 100) / 100
    
    akValue = AD_val - aeValue - ahValue - ajValue
    alValueNum = 0.18 if akValue > 2499 else 0.05
    amValue = akValue * alValueNum
    return akValue + amValue

def solveForAD(targetAN, discount, ajio):
    if targetAN <= 0: return 0
    
    low = 0
    high = targetAN * 10
    maxIters = 100
    mid = 0
    
    for i in range(maxIters):
        mid = (low + high) / 2
        currentAN = calculateAN(mid, discount, ajio)
        
        if abs(currentAN - targetAN) < 0.001:
            break
        
        if currentAN < targetAN:
            low = mid
        else:
            high = mid
    return mid

target = 756
discount = 65
ajio = 34

result = solveForAD(target, discount, ajio)
roundedResult = round(result * 100) / 100
print(f'Result: {result}')
print(f'Rounded Result (AE): {roundedResult}')
print(f'Final AO: {calculateAN(roundedResult, discount, ajio)}')
