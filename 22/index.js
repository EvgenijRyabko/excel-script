const xlsx = require("xlsx");
const excel = require("excel4node");
const fs = require("fs");

const parseData = (data) => {
  return data.map((el) => {
    delete el["__EMPTY"];

    return {
      name: el[Object.keys(el)[0]],
      nomNumber: el["__EMPTY_1"] ? el["__EMPTY_1"] : "",
      units: el["__EMPTY_2"] ? el["__EMPTY_2"] : "",
      priceInUAH: el["__EMPTY_3"] ? el["__EMPTY_3"] : 0,
      priceInRUB: el["__EMPTY_4"] ? el["__EMPTY_4"] : 0,
      restForMonthBegin_Amount: el["__EMPTY_5"] ? el["__EMPTY_5"] : 0,
      restForMonthBegin_Price: el["__EMPTY_6"] ? el["__EMPTY_6"] : 0,
      turnover_Received_Amount: el["__EMPTY_7"] ? el["__EMPTY_7"] : 0,
      turnover_Received_Price: el["__EMPTY_8"] ? el["__EMPTY_8"] : 0,
      turnover_Used_Amount: el["__EMPTY_9"] ? el["__EMPTY_9"] : 0,
      turnover_Used_Price: el["__EMPTY_10"] ? el["__EMPTY_10"] : 0,
      restForMonthEnd_Amount: el["__EMPTY_11"] ? el["__EMPTY_11"] : 0,
      restForMonthEnd_Price: el["__EMPTY_12"] ? el["__EMPTY_12"] : 0,
    };
  });
};

const groupDataByFIO = (data, startIndex) => {
  const result = [];

  for (let i = startIndex; i < data.length; i++) {
    if (data[i].name.includes("МОЛ")) {
      result.push({
        FIO: data[i].name.split(": ")[1],
        data: [],
      });
    } else {
      result[result.length - 1].data.push(data[i]);
    }
  }

  return result;
};

const parseFile = async (file) => {
  const workbook = xlsx.readFile(file);
  const sheet_name_list = workbook.SheetNames;
  return xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[5]]);
};

const writeFile = async (FIO, data) => {
  const workbook = new excel.Workbook({
    defaultFont: {
      name: "Arial",
      size: 10,
    },
    numberFormat: "$#,##0.00; ($#,##0.00); -",
  });
  const ws = workbook.addWorksheet("Sheet 1");
  const centeredStyle = workbook.createStyle({
    alignment: { vertical: "center", horizontal: "center" },
  });

  ws.cell(1, 1).string('ФГБОУ ВО "ЛГПУ"').style(centeredStyle);

  ws.cell(3, 1).string("Ведомость остатков ТМЦ:").style(centeredStyle);
  ws.cell(4, 1).string("по счету 20").style(centeredStyle);
  ws.cell(5, 1).string("на дату 01.06.23").style(centeredStyle);

  ws.cell(7, 1).string("МОЛ:").style(centeredStyle);

  ws.cell(7, 3)
    .string(
      `
    "${FIO}
МПК ФГБОУ ВО ""ЛГПУ""
не указано"
    `
    )
    .style(centeredStyle);

  ws.row(7).setHeight(36);
  ws.row(9).setHeight(50);
  ws.column(1).setWidth(41.14);
  ws.column(2).setWidth(9.71);
  ws.column(3).setWidth(24.86);
  ws.column(5).setWidth(8.86);

  ws.cell(9, 1).string("ТМЦ").style(centeredStyle);
  ws.cell(9, 2).string("КЕКР").style(centeredStyle);
  ws.cell(9, 3).string("Ед. изм.").style(centeredStyle);
  ws.cell(9, 4).string("Цена").style(centeredStyle);
  ws.cell(9, 5).string("Счет учета").style(centeredStyle);
  ws.cell(9, 6).string("Остатки").style(centeredStyle);
  ws.cell(9, 8)
    .string("Дата приобретения")
    .style({ alignment: { ...centeredStyle.alignment, wrapText: true } });
  ws.cell(10, 6).string("Кол-во").style(centeredStyle);
  ws.cell(10, 7).string("Сумма").style(centeredStyle);

  for (let i = 0; i < data.length; i++) {
    ws.cell(11 + i, 1).string(data[i].name);
    ws.cell(11 + i, 2).string("2210");
    ws.cell(11 + i, 3).string(data[i].units);
    ws.cell(11 + i, 4).number(data[i]?.priceInRUB || 0);
    ws.cell(11 + i, 5).string("221");
    ws.cell(11 + i, 6).number(data[i].restForMonthEnd_Amount);
    ws.cell(11 + i, 7).number(data[i].restForMonthEnd_Price);
    ws.cell(11 + i, 8).string("");
  }

  workbook.write(`${FIO}.xlsx`);
};

const start = async () => {
  const data = await parseFile("data.xlsx");

  const parsedData = parseData(data);

  const startIndex = parsedData.findIndex((el) => el.name.includes("МОЛ"));

  const groupedData = groupDataByFIO(parsedData, startIndex);

  for (let i = 0; i < groupedData.length; i++) {
    await writeFile(groupedData[i].FIO, groupedData[i].data);
  }
};

start();
