
// Test Cases
///console.log(isBalancedParenthesis("((())"));

let s='((())'

let isBalancedParenthesis = (str) => {
    
    return str.split('').reduce((uptoPrevChar, thisChar) => {
        if(thisChar === '(') {
            return --uptoPrevChar;
        } else if (thisChar === ')' ) {
            return ++uptoPrevChar;
        }

        return uptoPrevChar
    }, 0);
}

console.log(Math.abs(isBalancedParenthesis("((()"))); 


