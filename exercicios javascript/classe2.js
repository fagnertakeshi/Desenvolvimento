class Avo {
    constructor(sobrenome){
        this.sobrenome=sobrenome
    }
}

class Pai extends Avo {
    constructor(sobrenome= 'Mituyassu',profissao='Professor'){
        super(sobrenome)
        this.profissao=profissao
       // this.sobrenome=sobrenome
    }
  
}

class Filho extends Pai {
    constructor(){
        super('Silva')
    }
}

const filho  = new Filho

console.log(filho)