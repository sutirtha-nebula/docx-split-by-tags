import Color from "color";

const hex = Color("yellow").hex().replace("#", "");
console.log(hex); // "0000FF"
