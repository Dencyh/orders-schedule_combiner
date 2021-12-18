import XLSX from 'xlsx'
import fs from 'fs'
import chalk from 'chalk'

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
const wbOrders = XLSX.readFile(`${folderPath}/${ordersFileName}`, { cellDates: true })
const wsOrders = wbOrders.Sheets['отчет']

const ordersData = XLSX.utils.sheet_to_json(wsOrders)

// Getting schedules data
const
    wbSchedules = XLSX.readFile(`${folderPath}/${schedulesFileName}`, { cellDates: true }),
    wsSchedules = wbSchedules.Sheets['Sheet1'],
    schedulesData = XLSX.utils.sheet_to_json(wsSchedules)


// Getting unique couriers
let couriersFromShifts = []


ordersData.forEach((orderEntry) => {
    couriersFromShifts.push(orderEntry['Курьер'])
})
couriersFromShifts = Array.from(new Set([...couriersFromShifts])) // Unique couriers from shifts


const couriersDropships = []
const couriersDropshipsReturn = []

schedulesData.forEach((courier) => {
    if (courier['Тип транспортного средства'] == 'Дзерж Дропофф(только забор)') {
        couriersDropships.push([courier['Курьер']])
    } else if (courier['Тип транспортного средства'] == 'Возврат Дзерж КГТ') {
        couriersDropshipsReturn.push([courier['Курьер']])
    }
})



couriersDropships.forEach((courier) => {
    ordersData.forEach((orderInfo) => {
        if (
            courier[0] == orderInfo['Курьер']
            && (orderInfo['Статус'] == 'В процессе'
                || orderInfo['Статус'] == 'Выполнено')
        ) {
            courier.push(orderInfo['Адрес'])
            courier.push('')
        }
    })
})

couriersDropshipsReturn.forEach((courier) => {
    ordersData.forEach((orderInfo) => {
        if (
            courier[0] == orderInfo['Курьер']
            && (orderInfo['Статус'] == 'В процессе'
                || orderInfo['Статус'] == 'Выполнено')
        ) {
            courier.push(orderInfo['Адрес'])
            courier.push('')
        }
    })
})


// Sort by courier (first element) alphabetically
const couriersSortedByName = couriersDropships.sort((a, b) => {
    if (a[0] < b[0]) {
        return -1
    }
}).filter((element) => { //gets rid of the couriers that are not in the schedule and have 0 orders
    if (element.length < 2) console.log(chalk.green('Нет заказов по прямому потоку : ' + element))
    return element.length > 1
})

const couriersReturnSortedByName = couriersDropshipsReturn.sort((a, b) => {
    if (a[0] < b[0]) {
        return -1
    }
}).filter((element) => { //gets rid of the couriers that are not in the schedule and have 0 orders
    if (element.length < 2) console.log(chalk.green('Нет заказов по возвратному потоку: ' + element))
    return element.length > 1
})







// Get an array of couriers from schedules and compare them to those, who have orders. Console.log missing
const couriersFromSchedule = []
schedulesData.forEach((sheduleInfoRow) => {
    couriersFromSchedule.push(sheduleInfoRow['Курьер'])
})

couriersFromSchedule.forEach((courier) => {
    if (couriersFromShifts.indexOf(courier) < 0) console.log(chalk.red('Не получил маршрут (заказы переназначены) : ' + courier))
})

couriersFromShifts.forEach((courier) => {
    if (couriersFromSchedule.indexOf(courier) < 0) console.log(chalk.magenta('Убрали из расписания: ' + courier))
})



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

couriersReturnSortedByName.forEach((courier) => {
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

const couriersReturnSortedByCompany = couriersReturnSortedByName.sort((a, b) => {
    if (a[0] < b[0]) {
        return -1
    }
})

// Insert header to the result worksheet
couriersSortedByCompany.unshift(['', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])


couriersReturnSortedByCompany.unshift(['', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])


// Creating new file and writing to it
const dropshipsWS_name = 'Прямой поток'
const dropshipsWS = XLSX.utils.aoa_to_sheet(couriersSortedByCompany)
const dropshipsReturnWS_name = 'Возвратный поток'
const dropshipsReturnWS = XLSX.utils.aoa_to_sheet(couriersReturnSortedByCompany)
const newWB = XLSX.utils.book_new()

XLSX.utils.book_append_sheet(newWB, dropshipsWS, dropshipsWS_name)
XLSX.utils.book_append_sheet(newWB, dropshipsReturnWS, dropshipsReturnWS_name)


XLSX.writeFile(newWB, './ResultForm.xlsx')


