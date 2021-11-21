const XLSX = require('xlsx')
const fs = require('fs')

// Get file names
const
    folderPath = './Schedules_Orders',
    filesList = []

let schedules = new RegExp(/schedules*/)
let orders = new RegExp(/orders*/)

let
    schedulesFileName,
    ordersFileName


fs.readdirSync(folderPath).forEach(file => {
    filesList.push(file)
});

filesList.forEach((fileName) => {
    if (schedules.test(fileName)) {
        schedulesFileName = fileName
    } else if (orders.test(fileName)) {
        ordersFileName = fileName
    }
})

// Get orders data
const
    wbOrders = XLSX.readFile(`${folderPath}/${ordersFileName}`, { cellDates: true }),
    wsOrders = wbOrders.Sheets['отчет'],
    ordersData = XLSX.utils.sheet_to_json(wsOrders)


// Getting unique couriers
const couriers = []

ordersData.forEach((item) => {
    if (!couriers.includes(item['Курьер'])) {

        couriers.push(item['Курьер'])
    }
})

// Convert each courier value into separate array
const couriersUnique = []

couriers.forEach((item) => {
    if (!couriersUnique.includes(item)) {
        couriersUnique.push([item])
    }
})


couriersUnique.forEach((courier) => {
    ordersData.forEach((orderInfo) => {
        if (courier[0] == orderInfo['Курьер'] && (orderInfo['Статус'] == 'В процессе' || /* temp */ orderInfo['Статус'] == 'Выполнено')) {
            courier.push(orderInfo['Адрес'])
            courier.push('')
        }
    })
})
// Sort by courier (first element) alphabetically
const couriersSortedByName = couriersUnique.sort((a, b) => {
    if (a[0] < b[0]) {
        return -1
    }
}).filter((element) => { //gets rid of the couriers that are not in the schedule and have 0 orders
    if (element.length < 2) console.log(element)
    return element.length > 1
})





// Getting schedules data
const
    wbSchedules = XLSX.readFile(`${folderPath}/${schedulesFileName}`, { cellDates: true }),
    wsSchedules = wbSchedules.Sheets['Sheet1'],
    schedulesData = XLSX.utils.sheet_to_json(wsSchedules)

// Append couriers' company name to the beggining of the row
couriersSortedByName.forEach((courier) => {
    schedulesData.forEach((schedule) => {
        if (schedule['Курьер'] == courier[0]) {
            const ordersCount = Math.round((courier.length - 1) / 2)
            courier.splice(1, 0, schedule['Номер машины'], ordersCount)
            courier.unshift(schedule['Компания'])
        }
    })
})


// Sort by company (!first element) alphabetically
const couriersSortedByCompany = couriersSortedByName.sort((a, b) => {
    if (a[0] < b[0]) {
        return -1
    }
})


// Insert header to the result worksheet
couriersSortedByCompany.unshift(['', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])



// Creating new file and writing to it
const newWS_name = 'Results'

const newWS = XLSX.utils.aoa_to_sheet(couriersSortedByCompany)

const newWB = XLSX.utils.book_new()

XLSX.utils.book_append_sheet(newWB, newWS, newWS_name)

XLSX.writeFile(newWB, './EzForm.xlsx')


