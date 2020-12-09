/*Mr. Frefzie has a funny job just match if two person are blood related or not. He always try first with their names to see if they have same family name (last name) or not.
Please help him with this job.
Input:
Has two lines, each contains a name, consists of a first name and a last name and seperated by a space like this:
Rosanne Louser
Inell Hovi
Output
Different
*/
var names = input.split("\n")
var lastname1 = names[0].split(" ")
var lastname2 = names[1].split(" ")


if (lastname1[1]==lastname2[1])
      console.log('Same')
else
      console.log('Different')