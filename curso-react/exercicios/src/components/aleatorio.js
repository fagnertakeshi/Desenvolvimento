import React from 'react'
import {Text} from 'react-native'
import Estilo from './estilo'


export default (props) => {

    const num_ale=parseInt(Math.random() * (props.max - props.min) + props.min);
    console.warn(props)
    return (
    <Text style={Estilo.fontAl}> 
                     O valor aleatorio é {num_ale}
                    </Text>

    )

}