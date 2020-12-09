

values=[1,2,3,4,5]

var squarefuncwithadjustment = function(x, adjustment) {
    return (x*x + adjustment);
}

console.log(values.map(
    function(x) { return squarefuncwithadjustment(x, 2); }
  ))

