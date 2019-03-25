const path = require('path')
const fs = require('mz/fs')
const glob = require('glob')
const xlsx = require('node-xlsx');
const curPath = path.resolve(__dirname, './imgs')
let options = {
    cwd: curPath, // 在pages目录里找
    sync: true // 这里不能异步，只能同步
}
let globInstance = new glob.Glob('!(_)*', options)
let imgNameArr = globInstance.found
let imgNameArrLength = imgNameArr.length
let dateArr = [
    '2018/12/03',
    '2018/12/04',
    '2018/12/05',
    '2018/12/06',
    '2018/12/07'
]
let startStopName = [ //起止时间名字
    '12.03',
    '12.07'
]
let option = {
    id: ' ',
    date: ' ',//接单日期
    danwei: '招商银行信用卡中心',
    xuqiufang: '朱丹',
    xuqiuming: ' ',//需求名称
    type : ['banner','grid','开屏'],
    num : 1,
    price: ' ',
    designer: '王玲'
}
let excelData = [
    [
        'ID',
        '接单日期',
        '发包单位',
        '需求方',
        '需求名称',
        '类型',
        '数量',
        '价格',
        '设计师'
    ]
]
let heji = [
    '合计',
    ' ',
    ' ',
    ' ',
    ' ',
    ' ',
    imgNameArrLength,
    ' ',
    ' '
] //合计
let dateArrLength = dateArr.length
let remainder = imgNameArrLength % dateArrLength  //余数
let average = (imgNameArrLength - remainder) / dateArrLength //平均数
let LastNum = average + remainder //最后一位日统计数字
let newDateArr = []
dateArr.forEach(function(item,index) {
    if(index+1 === dateArr.length){
        for(let i = 0; i<LastNum ; i++){
            newDateArr.push(item)
        }
        newDateArr.push('最后日统计')
    } else {
        for(let i = 0; i<average ; i++){
            newDateArr.push(item)
        }
        newDateArr.push('日统计')
    }
})
let index = 0
let randomNum
//处理文件名不加后缀
imgNameArr = imgNameArr.map(function(item) {
    return item.substring(0, item.lastIndexOf('.'))
})
newDateArr.forEach(function(item, i) {
    let rowArr = [ //带图片文件名的每一行内容
        ' ', //id
        ' ',
        option.danwei, //单位
        option.xuqiufang, //xuqiufang
        ' ',//需求名称
        ' ',//type
        option.num, //num
        ' ', //price
        option.designer //designer
    ]
    if(item === '日统计'){
        excelData.push([ //日统计
            '日统计',
            ' ',
            ' ',
            ' ',
            ' ',
            ' ',
            average,
            ' ',
            ' '
        ])
    } else if(item === '最后日统计'){
        excelData.push([ //日统计
            '日统计',
            ' ',
            ' ',
            ' ',
            ' ',
            ' ',
            LastNum,
            ' ',
            ' '
        ])
    }else{
        randomNum = parseInt(3 * Math.random()) //0-3随机数随机type类型
        rowArr[1] = item
        rowArr[5] = option.type[randomNum]
        rowArr[4] = imgNameArr[index]
        excelData.push(rowArr)
        index++
    }
})
excelData.push(heji)
const range = {s: {c: 0, r:0 }, e: {c:0, r:3}}; // A1:A4
let option1 = {'!merges': [ range ]};
let buffer = xlsx.build([
    {
        name:'sheet1',
        data:excelData
    }
])

//将文件内容插入新的文件中
fs.writeFile(`王玲周报${startStopName[0]}-${startStopName[1]}.xlsx`,buffer,{'flag':'w'});


