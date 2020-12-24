function main (input) {
 input.split ("\n").forEach (function (n){
     console.log (n);
});
}
process.stdin.resume ();
process.stdin.setEncoding ("utf-8")
process.stdin.on ("data", function (input){
        stdin_input += input;
});
process.stdin.on ("end", function (){
      main (stdin_input);
});