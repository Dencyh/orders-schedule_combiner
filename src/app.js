const XLSX = require('xlsx')
const fs = require('fs')


const FOLDER_PATH = './Schedules_Orders'



// Get file names
const files = getFileNames(FOLDER_PATH)

// Get orders data
const ordersData = getData(files.orders, 'отчет')
const schedulesData = getData(files.schedule, 'Sheet1')


// Getting unique couriers
let couriersFromShifts = getCouriesList(ordersData)



const couriersDropships = []
const couriersDropshipsReturn = []

/* devide delivery and return */
schedulesData.forEach((courier) => devideByType(courier))

/* match address by name */
/* delivery */
couriersDropships.forEach(courier => matchAddress(courier))
/* return */
couriersDropshipsReturn.forEach((courier) => matchAddress(courier))



const dropped = []
/* delete dropped from the list or those  */
const couriersSortedByName = findDropped(couriersDropships).sort(sortByChar)


const droppedReturn = []
const couriersReturnSortedByName = findDropped(couriersDropshipsReturn).sort(sortByChar)



// Get an array of couriers from schedules and compare them to those, who have orders. Console.log missing
const couriersFromSchedule = []
schedulesData.forEach((sheduleInfoRow) => {
    couriersFromSchedule.push(sheduleInfoRow['Курьер'])
})


couriersFromShifts.forEach((courier) => {
    if (couriersFromSchedule.indexOf(courier) < 0) console.log('Убрали из расписания: ' + courier)
})



// Get info for active
couriersSortedByName.forEach((courier) => insertFullInfo(courier))
couriersReturnSortedByName.forEach((courier) => insertFullInfo(courier))

// Get info for dropped
dropped.forEach(courier => insertCompanyOnly(courier))
droppedReturn.forEach(courier => insertCompanyOnly(courier))


// Sort by company
const couriersSortedByCompany = couriersSortedByName.sort(sortByChar)
const couriersReturnSortedByCompany = couriersReturnSortedByName.sort(sortByChar)

// Insert header to the result worksheet
couriersSortedByCompany.unshift(['', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])
couriersReturnSortedByCompany.unshift(['', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])


/* class Sheet {
    constructor(name, data, cols = [{}]) {
        this.name = name
        this.sheet = XLSX.utils.aoa_to_sheet(data)
        this.sheet['!cols'] = cols // Can you do that?
    }
}

function createFile(sheetsArr, fileName) {
    const wb = XLSX.utils.book_new()


    sheetsArr.forEach((sheet) => {
        XLSX.utils.book_append_sheet(wb, sheet.data, sheet.name)
    })

    XLSX.utils.book_append_sheet(wb, ws, wsName)


    XLSX.writeFile(wb, `./${fileName}.xlsx`)
} */



// Helper functions

function getFileNames(folder) {
    const files = {}
    const filesList = []

    let schedules = new RegExp(/schedules*/)
    let orders = new RegExp(/orders*/)


    fs.readdirSync(folder).forEach(file => {
        filesList.push(file)
    });

    filesList.forEach((file) => {
        if (schedules.test(file)) {
            files.schedule = file
        } else if (orders.test(file)) {
            files.orders = file
        }
    })

    return files
}

function getData(fileName, sheetName) {
    const sheet = XLSX
        .readFile(`${FOLDER_PATH}/${fileName}`, { cellDates: true })
        .Sheets[sheetName]

    const data = XLSX.utils.sheet_to_json(sheet)

    return data
}

/* This is the only way not to miss anyone if they were removed from schedule by mistake */
function getCouriesList(ordersData) {
    let couriers = []
    ordersData.forEach((orderEntry) => {
        couriers.push(orderEntry['Курьер'])
    })
    couriers = Array.from(new Set([...couriers]))
    return couriers
}

function matchAddress(courier) {
    ordersData.forEach(orderInfo => {
        if (
            courier[0] == orderInfo['Курьер']
            && (orderInfo['Статус'] == 'В процессе' || orderInfo['Статус'] == 'Выполнено')
        ) {
            courier.push(orderInfo['Адрес'])
            courier.push('')
        }
    })
}

function devideByType(courier) {
    if (courier['Тип транспортного средства'] == 'Дзерж Дропофф(только забор)') {
        couriersDropships.push([courier['Курьер']])
    } else if (courier['Тип транспортного средства'] == 'Возврат Дзерж КГТ') {
        couriersDropshipsReturn.push([courier['Курьер']])
    }
}

function findDropped(couriers) {
    return (
        couriers.filter((element) => {
            if (element.length < 2) {
                dropped.push([element])
            }
            return element.length > 1
        })
    )
}

function insertFullInfo(courier) {

    schedulesData.forEach((schedule) => {
        if (schedule['Курьер'] == courier[0]) {
            const ordersCount = Math.round((courier.length - 1) / 2)
            courier.splice(1, 0, schedule['Номер машины'], ordersCount)
            courier.unshift(schedule['Компания'])
        }
    })

}

function insertCompanyOnly(courier) {
    schedulesData.forEach((schedule) => {
        if (schedule['Курьер'] == courier[0]) {
            courier.push(schedule['Компания'])
        }
    })
}

function sortByChar(a, b) {
    if (a[0] < b[0]) {
        return -1
    }
}


// Creating new file and writing to it
const newWB = XLSX.utils.book_new()

const dropshipsWS_name = 'Прямой поток'
const dropshipsWS = XLSX.utils.aoa_to_sheet(couriersSortedByCompany)
XLSX.utils.book_append_sheet(newWB, dropshipsWS, dropshipsWS_name)

const dropshipsReturnWS_name = 'Возвратный поток'
const dropshipsReturnWS = XLSX.utils.aoa_to_sheet(couriersReturnSortedByCompany)
XLSX.utils.book_append_sheet(newWB, dropshipsReturnWS, dropshipsReturnWS_name)

if (dropped.length > 0) {
    const droppedWS_name = 'Дропнуло Прямой поток'
    const droppedWS = XLSX.utils.aoa_to_sheet(dropped)
    droppedWS['!cols'] = [{ wpx: 170 }, { wpx: 115 }]
    XLSX.utils.book_append_sheet(newWB, droppedWS, droppedWS_name)
}

if (droppedReturn > 0) {
    const droppedReturnWS_name = 'Дропнуло Возвратный поток'
    const droppedReturnWS = XLSX.utils.aoa_to_sheet(droppedReturn)
    droppedReturnWS['!cols'] = [{ wpx: 170 }, { wpx: 115 }]
    XLSX.utils.book_append_sheet(newWB, droppedReturnWS, droppedReturnWS_name)
}

XLSX.writeFile(newWB, './ResultForm.xlsx')


