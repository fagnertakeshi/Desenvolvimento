
if (process.argv.length < 3) {
    console.log('Usage: node ' + process.argv[1] + ' FILENAME');
    process.exit(1);
  }
  // Read the file and print its contents.
  var fs = require('fs')
    , filename = process.argv[2];
  fs.readFile(filename, 'utf8', function(err, data) {
    if (err) throw err;
    console.log('OK: ' + filename);
    console.log(data)
    main()
  });

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