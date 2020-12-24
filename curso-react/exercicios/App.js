import { StatusBar } from 'expo-status-bar';
import React from 'react';
import { StyleSheet, Text, SafeAreaView ,View} from 'react-native';
// import FlexboxV1 from './src/components/layout/FlexboxV1';
import FlexboxV2 from './src/components/layout/FlexboxV2';
import Quadrado from './src/components/layout/Quadrado';
import Mega from './src/components/MegaSena.js/Mega';
import ListaProdutosV2 from './src/components/produtos/ListaProdutosV2';
// import ListaProdutos from './src/components/produtos/ListaProdutos';
//import Primeiro from './src/components/Primeiro'
//import MinMax from './src/components/MinMax'
//import Aleatorio from './src/components/aleatorio'
//import Titulo from './src/components/Titulo'
//import Botao from './src/components/Botao'
//import Pai from './src/components/indireta/Pai'
//import ContadorV2 from './src/components/contador/ContadorV2'
//import Diferenciar from './src/components/Diferenciar'
//import Parimpar from './src/components/Parimpar';
//import Familia from './src/components/relacao/Familia'
//import Membro from './src/components/relacao/Membro';
// import UsuarioLogado from './src/components/UsuarioLogado'
export default function App() {
  return (
    // <View style={styles.container}>
    <View>
      <Mega qtdNumeros={7}/>
      {/* <FlexboxV2/> */}
     {/* <ListaProdutosV2/>  */}
      {/* Double mustach - passando um objeto para o componente */}
{/*      <UsuarioLogado usuario={{nome:'Fagner',email:'fagnertakeshi@gmail.com'}} /> 
     <UsuarioLogado usuario={{nome:'',email:'fagnertakeshi@gmail.com'}} /> 
     <UsuarioLogado usuario={{nome:'Fagner',email:'fagnertakeshi@gmail.com'}} /> 
     <UsuarioLogado usuario={{nome:'Carlos',email:'Carlos@gmail.com'}} /> 
     <UsuarioLogado usuario={{nome:null,email:'Carlos@gmail.com'}} /> 
     <UsuarioLogado usuario={null} />  */}
     {/* Conhecendo a propriedade props.children
        d
       <Familia>
        <Membro nome="Fagner" sobrenome="Mituyassu" />
        <Membro nome="Fabricia" sobrenome="Mituyassu" />
      </Familia>
      <Familia>
        <Membro nome="Jorge" sobrenome="Nogueira" />
        <Membro nome="Jose" sobrenome="Nogueira" />
      </Familia>  */}
     {/* <Diferenciar/> */}
     {/* d<Parimpar num={5} /> */}
     {/* <ContadorV2></ContadorV2>  */}
      {/* <MinMax min="3" max="20"> </MinMax> */}
     {/*  <Aleatorio min="3" max="20"> </Aleatorio>
      <Titulo principal="Cadastro de produto"/>
      <Titulo secundario="Subtitulo de produto"/>
      <Primeiro/>
      <Botao/>
      <StatusBar style="auto" /> */}
    </View>
  );
}

const styles = StyleSheet.create({
  container: {
    flexGrow:1,  // pode crescer 
    backgroundColor: '#FFF', // cor de fundo vermelha
    flexDirection:'row',
    padding: 20
  },
});
