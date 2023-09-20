import * as excel from "excel4node";

export const writeFile = async (FIO, data, path) => {
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
    ws.cell(11 + i, 2)
      .string("2210")
      .style(centeredStyle);
    ws.cell(11 + i, 3)
      .string(data[i].units)
      .style(centeredStyle);
    ws.cell(11 + i, 4)
      .number(data[i].priceInRUB || 0)
      .style(centeredStyle);
    ws.cell(11 + i, 5)
      .string("221")
      .style(centeredStyle);
    ws.cell(11 + i, 6)
      .number(data[i].restForMonthEnd_Amount)
      .style(centeredStyle);
    ws.cell(11 + i, 7)
      .number(data[i].restForMonthEnd_Price)
      .style(centeredStyle);
    ws.cell(11 + i, 8).string("");
  }

  workbook.write(`${path}/${FIO}.xlsx`);
};
