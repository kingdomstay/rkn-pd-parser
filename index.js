import * as cheerio from 'cheerio';
import axios from 'axios';
import fs from 'fs';
import xl from 'excel4node';

// SELF_SIGNED_SSL_ERROR
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const outputData = [];

const loadTxtToArray = (filePath) => {
    try {
        const data = fs.readFileSync(filePath, 'utf8');
        return data.split('\n')
            .map(line => {
                const trimmed = line.trim();
                return trimmed.length === 0 ? 'SKIP': trimmed;
            })
    } catch (error) {
        throw new Error(`Ошибка при загрузке данных ИНН, убедитесь что существует файл ${filePath} в корне проекта`)
    }
}

const getHTMLData = async (inn) => {
    return axios.get(`https://pd.rkn.gov.ru/operators-registry/operators-list/?act=search&name_full=&inn=${inn}&regn=`);
}

const startParser = async (list) => {
    for (const inn of list) {
        const HTMLData = (await getHTMLData(inn)).data
        const $ = cheerio.load(HTMLData)

        const rootTableEl = $('#ResList1').find('tr.clmn1').find('td');
        const registerNumber = rootTableEl.first().text();
        const operatorName = rootTableEl.eq(1).find('a').first().text();
        const inclusionInRegistry = rootTableEl.eq(2).text();
        const registrationDate = rootTableEl.eq(3).text();
        const startDateOfProcessing = rootTableEl.eq(4).text();

        // Пропуск, нет данных или вывод реального номера
        const getNeededRegisterNumber = () => {
            if (inn === "SKIP") return ""
            if (registerNumber) return registerNumber
            return "НЕТ"
        }

        outputData.push({
            inn: inn === "SKIP" ? "" : inn,
            registerNumber: getNeededRegisterNumber(),
            operatorName,
            operatorType: registerNumber ? "юридическое лицо" : "",
            inclusionInRegistry,
            registrationDate,
            startDateOfProcessing,
        })
    }
}

const formatDataToXlsxAndSave = () => {
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('Результаты парсинга');

    // Заголовки
    ws.cell(1, 1)
        .string("ИНН")
    ws.cell(1, 2)
        .string("Номер в реестре операторов ПД")
    ws.cell(1, 3)
        .string("Наименование")
    ws.cell(1, 4)
        .string("Тип оператора")
    ws.cell(1, 5)
        .string("Основание включения в реестр")
    ws.cell(1, 6)
        .string("Дата регистрации уведомления")
    ws.cell(1, 7)
        .string("Дата начала обработки")

    let currentColumn = 2
    for (const data of outputData) {
        ws.cell(currentColumn, 1)
            .string(data.inn)
        ws.cell(currentColumn, 2)
            .string(data.registerNumber)
        ws.cell(currentColumn, 3)
            .string(data.operatorName)
        ws.cell(currentColumn, 4)
            .string(data.operatorType)
        ws.cell(currentColumn, 5)
            .string(data.inclusionInRegistry)
        ws.cell(currentColumn, 6)
            .string(data.registrationDate)
        ws.cell(currentColumn, 7)
            .string(data.startDateOfProcessing)

        currentColumn++
    }

    wb.write('Результаты.xlsx');
}

const bootstrap = async () => {
    console.log("Загрузка локальных данных")
    const innData = loadTxtToArray('inn.txt');
    console.log(`Локальные данные загружены (количество записей: ${innData.length})`)
    console.log(`Парсер обрабатывает данные...`)
    await startParser(innData)
    console.log(`Парсинг завершен успешно!`)
    console.log(`Преобразование в Excel формат...`)
    formatDataToXlsxAndSave();
    console.log(`Файл сохранен!`)
}





bootstrap();