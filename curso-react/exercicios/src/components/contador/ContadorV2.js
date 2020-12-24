import React , {useState} from 'react'
import {Text} from 'react-native'
import Estilo from '../estilo'
import ContadorDisplay from './ContadorDisplay'


export default _ => {

const [num,setNum]=useState(0)

return (
    <>
    <Text style={Estilo.txtG}>
        Contador V2
    </Text>
    <ContadorDisplay num={num}/>
    </>
)
}