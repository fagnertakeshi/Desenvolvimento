import React from 'react'
import {Button} from 'react-native'



export default (props) => {

    function executar() {
        console.warn('Exec!!!')
    }
    return (
        <Button title="Executar!"
        onPress={executar}//se passar executar() já executa ao ler a funcao
     />
    )

}
