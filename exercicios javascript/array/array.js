//console.log(typeof Array, typeof new Array, typeof [])
let aprovados =  new Array('Bia','Carlos','Ana')

//console.log(aprovados)

aprovados= ['Bia','Carlos','Ana']

aprovados.forEach(function(nome,indice){
    console.log(`${indice}) ${(nome)}`)
})

aprovados.forEach(nome => console.log(nome))

const exibirAprovados = aprovado => console.log(aprovado);

aprovados.forEach(exibirAprovados)


//console.log(aprovados[0])

//console.log(aprovados[1])

aprovados.push('Fagner')

//console.log(aprovados.length)

aprovados[9]='Takeshi'

//console.log(aprovados.length)

aprovados.sort()

//console.log(aprovados)


delete aprovados[1]

//console.log(aprovados)

aprovados.splice(2,1)

aprovados.splice(2,2,'Fagner1','Teste')


console.log(aprovados)
