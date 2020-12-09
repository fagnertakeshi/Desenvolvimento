  
  
function birthdayCakeCandles(candles) {

    let maximo= (a,b)=> Math.max(a,b)
    candles.sort();
    const max = candles.reduce(maximo)
    return candles.filter(candleHeight => candleHeight === max).length;
}

const candles = [2,2,3,4,4]
const result = birthdayCakeCandles(candles);

console.log(result)



