
Array.prototype.filter2 = function(callback){
    const newArray=[]
    for (let i=0;i<this.length;i++){
        if(callback(this[i],i,this))
            newArray.push(this[i])
    }
    return newArray

}

const produtos = [
    {nome:'Notebook', preco:2499,fragil:true},
    {nome:'Ipad Pro', preco:4199,fragil:true},
    {nome:'Copo de Vidro', preco:12.49,fragil:true},
    {nome:'Copo de plastico', preco:1899,fragil:false}
]

const verificafragil = p =>(p.fragil) 

const produtocaro = p =>p.preco>2400

/*
console.log(produtos.filter(function(p) {
    if (p.fragil)    
        return true
}))*/

console.log(produtos.filter2(verificafragil))


console.log(produtos.filter(produtocaro))