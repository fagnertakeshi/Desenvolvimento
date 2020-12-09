/*Sam's house has an apple tree and an orange tree that yield an abundance of fruit. Using the information given below, determine the number of apples and oranges that land on Sam's house.

In the diagram below:

The red region denotes the house, where  is the start point, and  is the endpoint. The apple tree is to the left of the house, and the orange tree is to its right.
Assume the trees are located on a single point, where the apple tree is at point , and the orange tree is at point .
When a fruit falls from its tree, it lands  units of distance from its tree of origin along the -axis. *A negative value of  means the fruit fell  units to the tree's left, and a positive value of  means it falls  units to the tree's right. *
*/


let s=7

let t=11

let a=5

let b=5

let m=3

let n=2

let apples = [-2,2,1]

let oranges = [5,-6]


function countApplesAndOranges(s, t, a, b, apples, oranges){

    /* função apples.map soma todos os elementos e cria um novo array a
    e a função filter coloca um limite entre os valores do array e depois o array final é contado
    */
const result_apple = apples.map(x=>x+a).filter(e =>(e>=s && e <=t)).length

const result_orange = oranges.map(x=>x+b).filter(e =>(e>=s && e <=t)).length

console.log(result_apple)

console.log(result_orange)   

}

let result = countApplesAndOranges(s, t, a, b, apples, oranges)


