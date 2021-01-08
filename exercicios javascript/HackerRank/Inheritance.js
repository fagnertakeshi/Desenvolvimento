class Person {
constructor(firstName, lastName, identification) {
    this.firstName = firstName;
    this.lastName = lastName;
    this.idNumber = identification;
}

printPerson() {
    console.log(
        "Name: " + this.lastName + ", " + this.firstName 
        + "\nID: " + this.idNumber
    )
    }
}

class Student extends Person {
/*	
*   Class Constructor
*   
*   @param firstName - A string denoting the Person's first name.
*   @param lastName - A string denoting the Person's last name.
*   @param id - An integer denoting the Person's ID number.
*   @param scores - An array of integers denoting the Person's test scores.
*/
// Write your constructor here

/*	
*   Method Name: calculate
*   @return A character denoting the grade.
*/
// Write your method here

constructor(firstName, lastName, identification,grades =[] ) {
    super(firstName, lastName, identification);
    this.grades = grades;
}

calculate(){
     
    const reduce_count=  this.grades.reduce((a, c,) => (a + c), 0)

    const result=reduce_count/this.grades.length


    switch(true) {
        case (result>=90 && result<=100):
          return "O"
          break;
        case (result>=80 && result<90):
            return "E"
          // code block
          break;
        case (result>=70 && result<80):
            return  "A"
           break; 
        case (result>=55 && result<70):
             return "P"
           break;
        case (result>=40 && result<55):
            return "D"
          break;
        case (result<40):
            return "T"
          break;      
      }
}

}



let student1 = new Student("Fagner","Takeshi","1234567",[80,100]);

console.log(student1.calculate())
