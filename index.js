var fs = require('fs');

let xlsx = require('node-xlsx');

let obj = xlsx.parse('./' + 'data.xls');

let amount = 0, // 补助金额
    carItem = [], // 车补
    mealItem = [], // 餐补
    allName = [], // 存放所有补助用户集合
    personnel = [], // 去重复的用户集合
    resultObj = [
        ['姓名', '车补', '餐补', '合计']
    ];

// 车补
let carSubsidize = obj[0].data.slice(3);
// 餐补
let mealSubsidize = obj[1].data.slice(2);
for (let i = 0; i < carSubsidize.length; i++) {
    if (carSubsidize[i][0] == undefined) continue;
    allName.push(carSubsidize[i][0])
};
personnel = Array.from(new Set(allName));
for (let i = 0; i < personnel.length; i++) {
    amount = 0;
    for (let y = 0; y < carSubsidize.length; y++) {
        if (personnel[i] == carSubsidize[y][0]) {
            amount += Number(carSubsidize[y][3]);
        }
    };
    carItem.push(personnel[i] + ':' + amount);
};


// 统计加班补助
amount = 0,
allName = [],
personnel = [];
for (let i = 0; i < mealSubsidize.length; i++) {
    if (mealSubsidize[i][5] == undefined) continue;
    allName = allName.concat(mealSubsidize[i][5].split(' '))
};
personnel = Array.from(new Set(allName));
for (let i = 0; i < personnel.length; i++) {
    amount = 0;
    if (personnel[i] == '') continue;
    for (let y = 0; y < allName.length; y++) {
        if (personnel[i] == allName[y]) {
            amount += 15;
        }
    };
    mealItem.push(personnel[i] + ':' + amount);
};

let departmentLotal = 0
for (let i = 0; i < mealItem.length; i++) {
    resultObj.push([
        mealItem[i].split(':')[0],
        0,
        mealItem[i].split(':')[1],
        mealItem[i].split(':')[1],
    ]);
    for (let y = 0; y < carItem.length; y++) {
        if (mealItem[i].split(':')[0] == carItem[y].split(':')[0]) {
            resultObj[i + 1][1] = carItem[y].split(':')[1];
            resultObj[i + 1][3] = Number(carItem[y].split(':')[1]) + Number(resultObj[i + 1][3]);
        }
    };
    departmentLotal += Number(resultObj[i + 1][3]);
}
resultObj.push(['部门合计',departmentLotal]);

let buffer = xlsx.build([{name: 'sheet1', data: resultObj}]);
fs.writeFileSync('./result.xlsx', buffer, {'flag':'w'});