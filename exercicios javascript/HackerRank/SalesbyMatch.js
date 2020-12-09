

const arr = [1,1,3,1,2,1,3,3,3,3]


var counts ={}

let sum

for (var i = 0; i < arr.length; i++) {
    var num = arr[i];
    counts[num] = counts[num] ? counts[num] + 1 : 1;
}
 
sum=0


console.log(counts)

for (const prop in counts) {
    if (counts[prop] %2 ==0)
        sum=sum + 1
  }

console.log(sum)