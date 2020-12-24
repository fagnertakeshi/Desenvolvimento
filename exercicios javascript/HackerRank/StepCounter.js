'use strict';

const fs = require('fs');
const assert = require('assert');
const readline = require('readline');

const OUTPUT_PATH="C:\Users\Usuário\Documents\GitHub\Desenvolvimento\exercicios javascript\HackerRank"


console.log(process.argv);

process.stdin.resume();
process.stdin.setEncoding('utf-8');

let inputString = '';
let currentLine = 0;

process.stdin.on('data', function(inputStdin) {
  inputString += inputStdin;
});

process.stdin.on('end', function() {
  inputString = inputString.split('\n');    
  main();
});

process.on('SIGINT', (code) => {
    console.log(`About to exit with code: ${code}`);
    inputString = inputString.split('\n');    
    main();
    process.exit()
  });

function readLine() {
  return inputString[currentLine++];
}

const counter = (function counter() {
  let value = 0;
  function changeBy(val) {
    value += val;
  }
 
  return {
    getValue: function() {
      return value;
    },
    increment: function() {
        changeBy(1);
      },
      decrement: function() {
        changeBy(-1);
      }
  }
})();
function getFixedCounter(k) {
  // write your code here
  return counter
}
function main() {
   
  //const ws = fs.createWriteStream(process.env.OUTPUT_PATH);

  const val = parseInt(readLine().trim());

  const c = getFixedCounter(val);
 // assert(!('changeBy' in c));

  const n = parseInt(readLine().trim());
  for (let i = 0; i < n; ++i) {
    const op = readLine().trim();
    if (op == '+') {
      c.increment();
    } else if (op == '-') {
      c.decrement();
    } else if (op === '?') {
      console.log(`${c.getValue()}\n`);
    }
  }
  
}
