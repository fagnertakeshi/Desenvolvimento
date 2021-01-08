
let a =[5,2,7,8,-2,25,25]
//let b= parseInt(a);

let arr = a.sort((a,b)=>a-b)

let n=arr.length

tempArr = [];

// To store modified array 
  
// Adding numbers from sorted array to  
// new array accordingly 
let ArrIndex = 0; 
let i,j;

// Traverse from begin and end simultaneously  
for (i = 0, j = n-1; (i <= n / 2 || j > n / 2); i++, j--) { 
   // console.log(arr[i])
   // console.log(arr[j])
    tempArr[ArrIndex] = arr[j]; 
    ArrIndex++; 
    tempArr[ArrIndex] = arr[i]; 
    ArrIndex++; 
} 

// Modifying original array 
for (let i = 0; i < n; i++) 
    arr[i] = tempArr[i]; 


console.log(tempArr)