const carType = new RegExp(/^Дзерж Дропофф/)

const string = 'Дзерж Дропофф(только забор) 5'

console.log(carType.test(string))