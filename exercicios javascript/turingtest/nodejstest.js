const dns = require('dns')
const { pathToFileURL } = require('url')

dns.lookup('127.0.0.1', function(err, result) {
  console.log(result)
})


dns.reverse('172.217.25.4', (err, hostnames) => {
    console.log(hostnames)
  })
  

  /*Node provides Buffer class which provides instances to store raw data similar to an array of integers but corresponds to a raw memory allocation outside the V8 heap.

Buffer class is a global class that can be accessed in an application without importing the buffer module.
*/


var buffer = Buffer.alloc(16)

buffer.write("Hello", "utf-8")

buffer.write(" world!", 5, "utf-8")

const user = {

    hello:'Welcoeme to Turing',

    sayWelcome() {
        console.log(this.hello);
    }
};

setTimeout(user.sayWelcome.bind(user),1000)

var path = require('path');

var x = path.normalize("c:/provas/...");

console.log(x);

