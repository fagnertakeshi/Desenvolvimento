


const roundfive = x=> ((Math.ceil(x / 5)*5-x) >= 3 || x<38) ?   x :Math.ceil(x / 5)*5

const result = grade.map(roundfive)


console.log(result)


const grade = [73,67,38,33]