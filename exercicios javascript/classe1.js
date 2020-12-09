class Lancamento{

    constructor(nome='Generico',valor=0,desconto=0) {
        this.nome=nome
        this.valor=valor
        this.desconto=desconto    

    }
}
class CicloFinanceiro{
    constructor(mes,ano){
        this.mes=mes
        this.ano=ano
        this.lancamentos=[]
    }

    addLancamento(...lancamentos){

        lancamentos.forEach(l=>this.lancamentos.push(l))
    }

    sumario() {

        let valorConsolidado=0
        this.lancamentos.forEach(l=>{
            valorConsolidado += l.valor + l.desconto
        })
        return valorConsolidado
    }

    
}


const salario=new Lancamento('Salario',1000,0)
const contaDeLuz = new Lancamento('Luz',-220,0)
const contaDeagua= new Lancamento('Agua', -329,0)

const contas = new CicloFinanceiro(6,2018)
contas.addLancamento(salario,contaDeagua,contaDeLuz)

console.log(contas.sumario())


