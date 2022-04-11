const XLSX = require('xlsx')
const fs = require('fs')


const SCHEDULES_ORDERS_PATH = './Schedules_Orders'
const VEHICLES_PATH = './Vehicles'



// Get file names
const schedulesOrders = getFileNames(SCHEDULES_ORDERS_PATH)
const vehicles = getFileNames(VEHICLES_PATH)


// Get orders data
const ordersData = getData(schedulesOrders.orders, 'отчет', SCHEDULES_ORDERS_PATH)
const schedulesData = getData(schedulesOrders.schedule, 'Sheet1', SCHEDULES_ORDERS_PATH)
const vehiclesData = getData(vehicles.vehicles, 'Sheet1', VEHICLES_PATH)




// Getting unique couriers
let couriersFromShifts = getCouriesList(ordersData)



const couriersDropships = []
const couriersDropshipsReturn = []
const couriersDelivery = []

/* devide delivery and return */
schedulesData.forEach((courier) => devideByType(courier))

/* match address by name */
/* intake */
couriersDropships.forEach(courier => matchAddress(courier))
/* return */
couriersDropshipsReturn.forEach((courier) => matchAddress(courier))
/* delivery */
couriersDelivery.forEach((courier) => matchAddress(courier))



const dropped = []
/* delete dropped from the list or those  */
const couriersSortedByName = findDropped(couriersDropships, dropped).sort(sortByChar)


const droppedReturn = []
const couriersReturnSortedByName = findDropped(couriersDropshipsReturn, droppedReturn).sort(sortByChar)

const droppedDelivery = []
const couriersDeliverySortedByName = findDropped(couriersDelivery, droppedDelivery).sort(sortByChar)



// Get an array of couriers from schedules and compare them to those, who have orders
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
couriersDeliverySortedByName.forEach((courier) => insertFullInfo(courier))




// Get info for dropped
dropped.forEach(courier => insertCompanyOnly(courier))
droppedReturn.forEach(courier => insertCompanyOnly(courier))
droppedDelivery.forEach(courier => insertCompanyOnly(courier))


// Sort by company
const couriersSortedByCompany = couriersSortedByName.sort(sortByChar)
const couriersReturnSortedByCompany = couriersReturnSortedByName.sort(sortByChar)
const couriersDeliverySortedByCompany = couriersDeliverySortedByName.sort(sortByChar)



findVehicle(couriersSortedByName)
// Insert header to the result worksheet
couriersSortedByCompany.unshift(['Объем ТС', 'Партнер', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])
couriersReturnSortedByCompany.unshift(['Партнер', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])
couriersDeliverySortedByCompany.unshift(['Партнер', 'ФИО', '', '', 'Адрес 1', '', 'Адрес 2', '', 'Адрес 3', '', 'Адрес 4', '', 'Адрес 5', '', 'Адрес 6', '', 'Адрес 7', '', 'Адрес 8', '', 'Адрес 9'])




// Helper functions

function getFileNames(folder) {
    const files = {}
    const filesList = []

    let schedules = new RegExp(/schedules*/)
    let orders = new RegExp(/orders*/)
    let orderType = new RegExp(/^TMM/)




    fs.readdirSync(folder).forEach(file => {
        filesList.push(file)
    });

    filesList.forEach((file) => {
        if (schedules.test(file)) {
            files.schedule = file
        } else if (orders.test(file)) {
            files.orders = file
        } else if (file == 'VehicleSize.xlsx') {
            files.vehicles = file
        }
    })

    return files
}

function getData(fileName, sheetName, folder) {
    const sheet = XLSX
        .readFile(`${folder}/${fileName}`, { cellDates: true })
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
    let orderType = new RegExp(/^TMM/)
    ordersData.forEach(orderInfo => {
        if (
            courier[0] == orderInfo['Курьер']
            && orderType.test(orderInfo['Номер заказа'])
            && (orderInfo['Статус'] == 'В процессе' || orderInfo['Статус'] == 'Выполнено')
        ) {
            courier.push(orderInfo['Адрес'])
            courier.push('')
        }
    })
}

function devideByType(courier) {
    const carType = new RegExp(/^Дзерж Дропофф/)

    if (carType.test(courier['Тип транспортного средства'])) {
        couriersDropships.push([courier['Курьер']])
    } else if (courier['Тип транспортного средства'] == 'Возврат Дзерж КГТ') {
        couriersDropshipsReturn.push([courier['Курьер']])
    } else if (courier['Тип транспортного средства'] == 'Ford Transit Дропофф+доставка') {
        couriersDelivery.push([courier['Курьер']])
    }
}

function findDropped(couriers, droppedArr) {

    return (
        couriers.filter((element) => {
            if (element.length < 2) {
                droppedArr.push([element])
            }
            return element.length > 1
        })
    )
}

function findVehicle(couriersArray) {
    couriersArray.forEach(courier => {

        const allVehicles = []
        vehiclesData.forEach(vehicle => {
            allVehicles.push(vehicle['Сцепка Имя и фамилия'])
        })

        if (allVehicles.includes(courier[1])) {
            vehiclesData.forEach(vehicle => {

                if (vehicle['Сцепка Имя и фамилия'] == courier[1]) {
                    courier.unshift(vehicle['Объем_1'])
                    return
                }

            })
        } else {
            courier.unshift('')
        }
    })
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

/* function matchVehicle() {
    ordersData.forEach(orderInfo => {
        if (courier[0] == )
    })
} */

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
dropshipsWS['!cols'] = [{ wpx: 60 }, { wpx: 80 }, { wpx: 140 }, { wpx: 75 }, { wpx: 20 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }, { wpx: 80 }, { wpx: 10 }]
XLSX.utils.book_append_sheet(newWB, dropshipsWS, dropshipsWS_name)

const dropshipsReturnWS_name = 'Возвратный поток'
const dropshipsReturnWS = XLSX.utils.aoa_to_sheet(couriersReturnSortedByCompany)
XLSX.utils.book_append_sheet(newWB, dropshipsReturnWS, dropshipsReturnWS_name)

const dropshipDelivery_name = 'Совмещенные'
const dropshipDeliveryWS = XLSX.utils.aoa_to_sheet(couriersDeliverySortedByCompany)
XLSX.utils.book_append_sheet(newWB, dropshipDeliveryWS, dropshipDelivery_name)

if (dropped.length > 0) {
    const droppedWS_name = 'Дропнуло Прямой поток'
    const droppedWS = XLSX.utils.aoa_to_sheet(dropped)
    droppedWS['!cols'] = [{ wpx: 170 }, { wpx: 115 }]
    XLSX.utils.book_append_sheet(newWB, droppedWS, droppedWS_name)
}

if (droppedReturn.length > 0) {
    const droppedReturnWS_name = 'Дропнуло Возвратный поток'
    const droppedReturnWS = XLSX.utils.aoa_to_sheet(droppedReturn)
    droppedReturnWS['!cols'] = [{ wpx: 170 }, { wpx: 115 }]
    XLSX.utils.book_append_sheet(newWB, droppedReturnWS, droppedReturnWS_name)
}

if (dropped.length > 0) {
    const droppedDeliveryWS_name = 'Дропнуло Совмещенные'
    const droppedDeliveryWS = XLSX.utils.aoa_to_sheet(droppedDelivery)
    droppedDeliveryWS['!cols'] = [{ wpx: 170 }, { wpx: 115 }]
    XLSX.utils.book_append_sheet(newWB, droppedDeliveryWS, droppedDeliveryWS_name)
}

XLSX.writeFile(newWB, './ResultForm.xlsx')


