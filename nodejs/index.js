const express =  require('express');

const server= express();

server.use(express.json());




//Query params = ?nome=NOdejs
// Route Params /curso/2
// Request Body = { nome}

const cursos = ['Node Js','JavaScript','React Native']


function checkCurso(req,res,next){
    if(!req.body.name){
        return res.status(400).json({ error: "Nome do curso é obrigatório"});
    }
    return next();

}

function checkIndexCurso(req,res,next){
    if(!cursos[req.body.index]){
        return res.status(400).json({ error: "Indice do curso é  obrigatório"});
    }
    return next();

}

server.get('/cursos', (req,res)=>{
    return res.json(cursos);
});

server.use((req,res,next)=>{
    console.log(`URL CHAMADA: ${req.url}`);

    return next();
    
});



//server.get('/curso', (req,res)=>{
server.get('/cursos/:index',checkIndexCurso, (req,res)=>{
    const {index} =req.params;    
    console.log('ACESSOU A ROTA')
   // const nome= req.query.nome
   //const id = req.params.id
   //return res.json({"curso": `Curso ${id}` })
   return res.json (cursos[index]);
});

//criando um curso
server.post('/cursos', checkCurso,(req,res)=>{
    const {name} = req.body;
    cursos.push(name);
    return res.json(cursos);

});



//editando um curso
server.put('/cursos/:index',checkIndexCurso,(req,res)=>{
    const {index} = req.params;
    const {name} =  req.body;
    cursos[index]=name;
    return res.json(cursos);
})

//DELETANDO UM CURSO
server.delete('/cursos/:index',checkIndexCurso,(req,res)=>{

    const {index} = req.params;
    
    cursos.splice(index,1)

    return res.json({message: 'Curso deletado com sucesso'});

});
server.listen(3000,function(){
    console.log("Servidor rodando");
});
