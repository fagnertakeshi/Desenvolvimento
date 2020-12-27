import { StatusBar } from 'expo-status-bar';
import React, { Component } from 'react';
import { Button, StyleSheet, Text, View } from 'react-native';
import Botao from './src/components/Botao';
import Display from './src/components/Display';

const initialState = {
  displayValue:'0',
  clearDisplay:false,
  operation:null,
  values:[0,0],
  current:0,
}

export default class App  extends Component{
  state = {...initialState}


  addDigit = n => {
    console.debug(typeof this.state.displayValue)

    const clearDisplay = this.state.displayValue==='0'
      || this.state.clearDisplay
    
    if (n==='.'&& !clearDisplay && this.state.displayValue.includes('.')){
       return
     }  
    const currentValue = clearDisplay ? '': this.state.displayValue
    const displayValue= currentValue + n 
    this.setState({displayValue, clearDisplay:false})
    if (n!== '.') {
      const newValue = parseFloat(displayValue)
      const values= [...this.state.values]
      values[this.state.current]= newValue
      this.setState({values})
    }
  }

  clearMemory = () => {
    this.setState({...initialState})

  }

  setOperation = operation => {
    if (this.state.current===0){
      this.setState({operation, current:1,clearDisplay:true})
    } else {
      const equals= operation ==='='
      const values =[...this.state.values]
      try {
        values[0]= eval(`${values[0]} ${this.state.operation} ${values[1]}`)
      } catch(e) {
        values[0]= this.state.values[0]
      }

      values[1]=0
      this.setState({
         displayValue:`${values[0]}`,
         operation: equals ? null:operation,
         current:equals? 0:1,
         clearDisplay:true,
         values,

      })
    }
  }


  render() {
  return (
    <View style={styles.container}>
      <Display value={this.state.displayValue}/> 
      <View style={styles.buttons}>
      <Botao label='AC'  triple onClick={this.clearMemory}/>
      <Botao label='/' operation onClick={() => this.setOperation('/')}/> 
      <Botao label='7' onClick={()=> this.addDigit(7)}/> 
       <Botao label='8' onClick={()=> this.addDigit(8)}/> 
       <Botao label='9' onClick={()=> this.addDigit(9)}/> 
       <Botao label='*' operation onClick={() => this.setOperation('*')}/> 
      <Botao label='4' onClick={()=> this.addDigit(4)}/> 
      <Botao label='5' onClick={()=> this.addDigit(5)}/> 
      <Botao label='6' onClick={()=> this.addDigit(6)}/> 
      <Botao label='-' operation onClick={() => this.setOperation('-')}/> 
      <Botao label='1' onClick={()=> this.addDigit(1)}/>
      <Botao label='2' onClick={()=> this.addDigit(2)}/> 
      <Botao label='3' onClick={()=> this.addDigit(3)}/> 
      <Botao label='+' operation onClick={() => this.setOperation('+')}/> 
      <Botao label='0' double onClick={()=> this.addDigit(0)}/> 
      <Botao label='.' onClick={() => this.addDigit('.')}/> 
      <Botao label='=' operation onClick={() => this.setOperation('=')}/> 
      </View>
    </View>
  );
  }
}

const styles = StyleSheet.create({
  container: {
    // flex: 1,
    backgroundColor: '#fff',
    width:'370px',
    height:'640px',
    flexWrap:'wrap',
    backgroundColor: "beige",
    borderWidth: 5,
  },
  buttons:{
    flexDirection:'row',
    flexWrap:'wrap',
  }

});
