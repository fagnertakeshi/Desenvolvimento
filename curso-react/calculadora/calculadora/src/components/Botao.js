import React from 'react'
import {
        StyleSheet,
        Text,
        Dimensions,
        TouchableHighlight
} from 'react-native'

const styles= StyleSheet.create ({
    button: {
        fontSize:40,
        // height: Dimensions.get('window').width/4,
        // Width: Dimensions.get('window').width/4,
        height: '20',
        width: '90px',
        padding:20,
        backgroundColor : '#f0f0f0',
        textAlign: 'center',
        borderWidth: 1,
        borderColor: '#888',
        alignItems: 'flex-start'
    } ,
    operationButton: {
        color:'#fff',
        backgroundColor:'#fa8231'
    },
    buttonDouble: {
        //  width:(Dimensions.get('window').width/4) * 2
         width:180
    },
    buttonTriple: {
        // width:(Dimensions.get('window').width/4) * 3
        width:270
   },
})

export default props => {
        const stylesButton=[styles.button]
        if (props.double)  stylesButton.push(styles.buttonDouble)
        if (props.triple) stylesButton.push(styles.buttonTriple)
        if (props.operation) stylesButton.push(styles.operationButton)

        return (
            <TouchableHighlight onPress={props.onClick}>
                <Text style={stylesButton}>{props.label}</Text>
            </TouchableHighlight>
        )
}