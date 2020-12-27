import React from 'react'
import {
        StyleSheet,
        Text,
        View
} from 'react-native'

const styles= StyleSheet.create ({
    display: {
        // flex:1,
        width:'360px',
        height:'150px',
        padding:20,
        backgroundColor : 'rgba(0,0,0,0.6)',
        justifyContent:'center',
        alignItems:'flex-end',
        textAlign: 'center',
        borderWidth: 1,
        borderColor: '#888',
    },
    displayValue:{
        fontSize:60,
        color:'#ffff'
    }
})

export default props=> {
    return (
    <View style={styles.display}>
            <Text style={styles.displayValue}
            numberOfLines={1}>{props.value}                
            </Text>
    </View>
    )

}

