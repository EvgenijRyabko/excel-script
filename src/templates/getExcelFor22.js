import * as excel from "excel4node";

export const getExcelFor22 = async (account, FIO, data, path) => {
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

  ws.cell(1, 1)
    .string('ФГБОУ ВО "ЛГПУ"')
    .style(centeredStyle)
    .style({ font: { bold: true } });

  ws.cell(3, 1, 3, 8, true)
    .string("Ведомость остатков ТМЦ:")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(4, 1, 4, 8, true)
    .string("по счету 22")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(5, 1, 5, 8, true)
    .string("на дату 01.06.23")
    .style(centeredStyle)
    .style({ font: { bold: true } });

  ws.cell(7, 1)
    .string("МОЛ:")
    .style(centeredStyle)
    .style({ font: { bold: true } });

  ws.cell(7, 2)
    .string(
      `${FIO}
  МПК ФГБОУ ВО "ЛГПУ"
  не указано`
    )
    .style({
      alignment: {
        horizontal: "left",
        wrapText: true,
      },
      font: { bold: true, underline: true },
    });

  ws.row(7).setHeight(36);
  ws.row(9).setHeight(50);
  ws.column(1).setWidth(41.14);
  ws.column(2).setWidth(24.86);
  ws.column(3).setWidth(15.86);
  ws.column(5).setWidth(8.86);

  ws.cell(9, 1)
    .string("ТМЦ")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(9, 2)
    .string("КЕКР")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(9, 3)
    .string("Ед. изм.")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(9, 4)
    .string("Цена")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(9, 5)
    .string("Счет учета")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(9, 6, 9, 7, true)
    .string("Остатки")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(9, 8)
    .string("Дата приобретения")
    .style({
      alignment: {
        ...centeredStyle.alignment,
        wrapText: true,
      },
      font: { bold: true },
    });
  ws.cell(10, 6)
    .string("Кол-во")
    .style(centeredStyle)
    .style({ font: { bold: true } });
  ws.cell(10, 7)
    .string("Сумма")
    .style(centeredStyle)
    .style({ font: { bold: true } });

  for (let i = 0; i < data.length; i++) {
    if (
      ["всего", "итого"].some((el) => data[i].name.toLowerCase().includes(el))
    ) {
      ws.cell(11 + i, 1)
        .string(data[i].name)
        .style({ font: { bold: true } });
      ws.cell(11 + i, 6)
        .number(data[i].restForMonthEnd_Amount || 0)
        .style(centeredStyle)
        .style({ font: { bold: true } });
      ws.cell(11 + i, 7)
        .number(data[i].restForMonthEnd_Price || 0)
        .style(centeredStyle)
        .style({ font: { bold: true } });
    } else if (data[i].name.toLowerCase().includes("спецсчет")) {
      ws.cell(11 + i, 1)
        .string(data[i].name)
        .style({ font: { bold: true } });
    } else {
      ws.cell(11 + i, 1).string(data[i].name);
      ws.cell(11 + i, 2)
        .string("2210")
        .style(centeredStyle);
      ws.cell(11 + i, 3)
        .string(data[i].units || "")
        .style(centeredStyle);
      ws.cell(11 + i, 4)
        .number(data[i].price || 0)
        .style(centeredStyle);
      ws.cell(11 + i, 5)
        .string(account)
        .style(centeredStyle);
      ws.cell(11 + i, 6)
        .number(data[i].restForMonthEnd_Amount || 0)
        .style(centeredStyle);
      ws.cell(11 + i, 7)
        .number(data[i].restForMonthEnd_Price || 0)
        .style(centeredStyle);
    }
    ws.cell(11 + i, 8).string("");
  }
  workbook.write(`${path}/${FIO}.xlsx`);
};
