import React from 'react'
import {Text, View} from 'react-native'
import Estilo from './estilo'
export default props => {
  return (
    <View>
        {/* Renderização Condicional */}
        {props.num %2 ==0
        ?<Text style={Estilo.txtG}>Par</Text>
        :<Text style={Estilo.txtG}>Impar</Text>
        }
    </View>
        )
}