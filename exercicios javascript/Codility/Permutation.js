

let a =[1,2]

function solution(a) {
    // write your code in JavaScript (Node.js 8.9.4)
    a.sort()
    for (let i = 0; i < a.length; i++) {
        if(a[i+1]-a[i]>1){
            return 0
        }
    }
    return 1
}



console.log(solution(a))


