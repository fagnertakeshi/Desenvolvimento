const nums = [3,1,2,4,3]

let diff

const somar = (a,b) => a+b 

diff=nums[0]

let sum = nums.slice(1, nums.length).reduce(somar)

let result = Math.abs(sum - diff)

for (i = 1; i < nums.length-1; i++) {
    sum=sum - nums[i]
    diff=diff+ nums[i]
    if (Math.abs(diff - sum) < result) {
        result = Math.abs(diff - sum);
    }

}

console.log(result)