import React from 'react'
import { Button, Text, TextInput, View} from 'react-native'
import estilo from '../estilo'
import Estilo from '../estilo'




export default class Mega extends React.Component {
    
   state = {
            qtdNumeros:this.props.qtdNumeros,
            numeros:[]
        }

    gerarNumeros = ()=> {
       const numeros =  Array(this.state.qtdNumeros)
            .fill()
            .reduce(n => [...n,this.gerarNumeroNaoContido(n)],[])
        this.setState({numeros})
    }

    gerarNumeroNaoContido = nums => {
        const novo= parseInt(Math.random() * 60) + 1
        return nums.includes(novo) ? this.gerarNumeroNaoContido(nums): novo
    }

    alterarQtdNumero =(qtde)=> {
        this.setState({qtdNumeros: +qtde})
    }
    
    render() {
        return (
            <>
            <Text style={Estilo.txtG}>
                Gerador de Mega-Sena {this.state.qtdNumeros}
            </Text>
            <TextInput
                keyboardType={'numeric'} 
                style={Estilo.inputstyle}
                placeholder="Qtde de Numeros"
                value={this.state.qtdNumeros}
                onChangeText={qtde=>this.alterarQtdNumero(qtde)}                
                />
                <Button
                    title='Gerar'
                    onPress={this.gerarNumeros}>
                </Button>
                <View >
                <Text style={Estilo.txtG}>
                    {this.state.numeros.join(',')}
                </Text>
                </View>
                
            </>
              

        )
    }

}

