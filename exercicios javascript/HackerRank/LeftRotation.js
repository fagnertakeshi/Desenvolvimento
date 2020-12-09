function CyclicRotation(vector,k) {
	
	var newvector=[];
		
	for (i = 0; i < vector.length; i++) {
		j= (i+k)%vector.length;
		newvector[j]=vector[i];
		
	}
    return newvector;
}
const candles = [2,2,3,4,4]

const k=3


var Rotation = function(e,i,array, k) {
    return array[(i+k)%array.length]
}


let resultado= candles.map(function(e,i,array) { return Rotation(e,i,array, k)})

