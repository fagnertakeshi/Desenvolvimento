


var array1 = '1 2 3 4 4'
var array2 = '2 3 1 3 3'

var values1 = array1.split(" ")
var values2 = array2.split(" ")

var val_sort1 = values1.sort()
var val_sort2 = values2.sort()

var sum=0
for (i = 0; i < values1.length; i++) {
    sum= sum + val_sort1[i]*val_sort2[i]
}


console.log(sum)