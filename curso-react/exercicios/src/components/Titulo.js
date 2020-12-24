import React from 'react';
import { StyleSheet, Text, View } from 'react-native';
import Estilo from './estilo'


export default(props) => {
    console.warn('Opa')
    return (
        <React.Fragment>
            <Text style={Estilo.txtG}>{props.principal} </Text>
            <Text>{props.secundario} </Text>
        </React.Fragment>
    )

}