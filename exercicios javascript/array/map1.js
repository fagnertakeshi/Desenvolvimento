const nums = [1,2,3,4,5]

let resultado= nums.map(function(e){
    return 2*e
})


const soma10 = e => e+10

const triplo = e => 3*e

const paraDinheiro = e => `R$ ${parseFloat(e).toFixed(2).replace('.',',')}`


resultado= nums.map(soma10).map(triplo).map(paraDinheiro)

console.log(resultado)
