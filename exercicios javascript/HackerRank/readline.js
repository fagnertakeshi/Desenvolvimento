const fs = require('fs');
const readline = require('readline');
const lines = require('fs').readFileSync('input.txt', 'utf-8') .split('\r\n')
console.log(lines)



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
     
 const val = lines[0];

const c = getFixedCounter(val);
  
const n = lines[1];


  for (let i = 0; i < lines.length; ++i) {
    const op =lines[i];
    if (op == '+') {
      c.increment();
    } else if (op == '-') {
      c.decrement();
    } else if (op === '?') {
      console.log(`${c.getValue()}\n`);
    }
  }
  

  }


  main()