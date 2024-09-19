"use strict";

var x;   //Variable Declaration
var x = undefined  //Variable Declaration
x = 10;
var y = 15; //Variable declaration with value assignment
var z = null;   //Non-existent or a invalid value

var symbol1 = Symbol('symbol'); //Anonymous and unique value

arr1 = ['abc', 'mno', '123', 6, true];

typeof(x);
typeof(arr1);  //to get the type of arr1 variable

var obj1 = {
    a: 3,
    b: "Hola",
    z: function(){
        return this.a;  //will return value 3
    }
}

//hoisting - behavoiur of JS where all variable & function delcarations are moved at top

//"==" to compare Values
//"===" to compare Values &also  Type

var x = 2;
var y = "2";

(x == 2);   //true  --as value is same
(x === 2);  //false   --as Diff types

//Implicit Type Coercion
(x + y) // 22 --> as one is string both turns into string

// var can be used anywehre
// let can be used only in the block where it is decleard

//JS is a loosely typed language

//In JS primitive data types are passed by value and non-primitive data types are passed by reference


Encapsulation - Hiding the data / Bundling of data & methods together
Polymorphism - Having many forms
    Compile time Polymorphism - static/early binding
    Run Time Polymorphism - dynamic/late binding

Data Abstraction - Show only necessary data
Inheritance - a class derived from another class

