var m1 = new Map();

m1.set("orange", 10);
m1.set("apple", 5);
m1.set("banana", 20);
m1.set("cherry", 13);

console.log(m1);

let m2= new Map([...m1.entries()].sort((a,b) => b[1] - a[1]));

console.log(m2);