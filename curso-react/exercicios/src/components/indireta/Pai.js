import React, {useState} from 'react'
import Filho from './Filho'
import {Text} from 'react-native'

export default props => {
    const [num,SetNum]= useState(0)

    function exibirValor(numero) {
        SetNum(numero)

    }
return (
    <>
    <Text>{num}</Text>
    <Filho 
        min={1}
        max= {10}
    funcao={exibirValor}
    />
    </>

    )
}