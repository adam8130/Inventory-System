const http = require('http')
const url = require('url')
const fs = require('fs')
const XLSX = require('xlsx')



// 定義變量
let tmpArr
let locationArr
let inventoryArr
let locationPath = '../Data/Location.xls'
let inventoryPath = '../Data/Inventory.xlsx'
const { encode_cell, decode_range } = XLSX.utils


// 讀取總部儲位
function readHqLocation () {
    console.log('HQ-Location. re-build at ', Date())

    locationArr = []
    const workbook = XLSX.readFile(locationPath)
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    
    const range = decode_range(sheet['!ref'])
    const startRow = range.s.r + 1
    const endRow = range.e.r
    
    for (let row = startRow; row <= endRow; row++) {
        const model = sheet[encode_cell({ r: row, c: 0 })].v
        const location = sheet[encode_cell({ r: row, c: 1 })].v
        locationArr.push({model: model, location: location})
    }
    console.log(locationArr.length)
} readHqLocation()


// 讀取全商品庫存
function readInventory () {
    console.log('Inventory. re-build at ', Date())

    tmpArr = []
    inventoryArr = []
    const workbook = XLSX.readFile(inventoryPath)
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    
    const range = decode_range(sheet['!ref'])
    const startRow = range.s.r + 1
    const endRow = range.e.r
    const endCol = range.e.c

    for (let row = startRow; row <= endRow; row++) {
        if (sheet[encode_cell({ r: row, c: 2 })]) {

            for(let col = 1; col <= endCol; col++) {
                tmpArr.push({
                    title: sheet[encode_cell({ r: 0, c: col })].v,
                    value: sheet[encode_cell({ r: row, c: col })].v
                })    
            }
            inventoryArr.push(tmpArr)
            tmpArr = []
        }
    }
    console.log(inventoryArr.length)
} readInventory()


// 設置防斗函數
function debounce(callback, delay) {
    let timer
    return function(...arg) {
        clearTimeout(timer)
        timer = setTimeout(() => callback.apply(this, arg), delay)
    }
}
const debouncedLocationEvent = debounce(() => {
    locationArr = null
    readHqLocation()
}, 5000)
const debouncedInventoryEvent = debounce(() => {
    inventoryArr = null
    readInventory()
}, 5000)


// 監聽儲位文件是否被更新
let prevLocationStat
fs.stat(locationPath, (err, stats) => {
    if (err) return
    prevLocationStat = stats.mtime
})
fs.watch(locationPath, (event) => {
    if (event == 'change') {
        console.log(event)
        fs.stat(locationPath, (err, stats) => {
            if (err) return
            if (stats.mtime !== prevLocationStat) 
            debouncedLocationEvent()
        })
    }
})


// 監聽庫存文件是否被更新
let prevInventoryStat
fs.stat(inventoryPath, (err, stats) => {
    if (err) return
    prevInventoryStat = stats.mtime
})
fs.watch(inventoryPath, (event) => {
    console.log(event)
    fs.stat(inventoryPath, (err, stats) => {
        if (err) return
        if (stats.mtime !== prevInventoryStat) 
        debouncedInventoryEvent()
    })
})


// 創建服務器
const server = http.createServer((req, res) => {

    const requestUrl = url.parse(req.url, true)
    const pathName = requestUrl.pathname
    const model = requestUrl.query.model

    console.log(pathName, model)
    if (pathName === '/api') {
        
        res.writeHead(200, {
            'content-type': 'application/json',
            "access-control-allow-origin": "*"
        })

        if (model.length !== 10) {
            const locationResult = locationArr.filter((item) => item.model === model)
            res.end(JSON.stringify({ locationResult }))
        }
        else {
            const locationResult = locationArr.filter((item) => item.model === model.substring(0, 7))
            const inventoryResult = inventoryArr.filter((item) => item[0].value == model)
            res.end(JSON.stringify({ locationResult, inventoryResult }))
        }
        
    }

})

server.listen(3000)