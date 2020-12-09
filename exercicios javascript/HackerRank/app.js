const http = require('http');

const hostname = '127.0.0.1';
const port = 3000;

const server = http.createServer((req, res) => {
  res.statusCode = 200;
  res.setHeader('Content-Type', 'text/plain');
  var ar=[10,11,12,13,14];
  let result = simpleArraySum(ar);
  alert(result);
  res.end(result);
  
});

server.listen(port, hostname, () => {
  console.log(`Server running at http://${hostname}:${port}/`);
});

function simpleArraySum(ar) {
    /*
     * Write your code here.
     */
    var i, sum;
    for (i = 0; i < ar.length; i++) {
        sum=sum + ar[i];   
    }
    return sum;
}
