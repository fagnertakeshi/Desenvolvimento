
let a =[1,2,1,4,5,3,1,4]
let x=5
let sum=0

function FrogRiver(x, a) {

let B = Array(x).fill(0)

for (let i = 0; i < a.length; i++) {
    if (a[i] <= x && B[a[i]-1] === 0){
        B[a[i]-1] = 1;
        sum += 1;
        if (sum === x){
            return i
        } 
    }
}

return (-1) 

}
console.log(FrogRiver(x,a))
