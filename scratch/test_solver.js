function calculateAN(AD_val, discount, ajio) {
  const aeValue = Math.round((AD_val * discount) / 100);
  const afValue = AD_val - aeValue;
  const agValueNum = afValue > 2499 ? 0.18 : 0.05;
  const ahValue = afValue - (afValue / (1 + agValueNum));
  const ajValue = Math.round(((afValue * ajio) / 100) * 100) / 100;
  
  const akValue = AD_val - aeValue - ahValue - ajValue;
  const alValueNum = akValue > 2499 ? 0.18 : 0.05;
  const amValue = akValue * alValueNum;
  return akValue + amValue;
}

function solveForAD(targetAN, discount, ajio) {
  if (targetAN <= 0) return 0;
  
  let low = 0;
  let high = targetAN * 10;
  let maxIters = 100;
  let mid = 0;
  
  for (let i = 0; i < maxIters; i++) {
    mid = (low + high) / 2;
    let currentAN = calculateAN(mid, discount, ajio);
    
    if (Math.abs(currentAN - targetAN) < 0.001) {
      break;
    }
    
    if (currentAN < targetAN) {
      low = mid;
    } else {
      high = mid;
    }
  }
  return mid;
}

const target = 756;
const discount = 65;
const ajio = 34;

const result = solveForAD(target, discount, ajio);
const roundedResult = Math.round(result * 100) / 100;
console.log('Result:', result);
console.log('Rounded Result (AE):', roundedResult);
console.log('Final AO:', calculateAN(roundedResult, discount, ajio));
