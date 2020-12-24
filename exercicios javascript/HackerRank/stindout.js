process.stdin.resume();
process.stdin.setEncoding('utf8');

process.stdin.on('data', function(chunk) {
  process.stdout.write('data: ' + chunk);
});

process.on('SIGINT', function(){
    process.stdout.write('\n end \n');
    process.exit();
});