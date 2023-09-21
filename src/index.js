import { checkDirectory } from "./utils/checkDir.js";
import { rl } from "./utils/readline.js";
import { parseData, groupDataByFIO } from "./utils/parsers.js";
import { findHeadersPlaces } from "./utils/findHeadersPlaces.js";
import templates from "./templates/index.js";
import * as xlsx from "xlsx/xlsx.mjs";
import * as fs from "fs";

const parseFile = async (file) => {
  const buf = fs.readFileSync(file);
  const workbook = xlsx.read(buf);
  const sheet_name_list = workbook.SheetNames;
  return xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[5]]);
};

const start = async () => {
  rl.question("Введите имя файла:\n", async (fileName) => {
    try {
      if (!fileName) {
        throw new Error("Ошибка ввода");
      }

      //* Get data from excel file
      const data = await parseFile(`filesToParse/${fileName}.xlsx`);

      const startIndex = data.findIndex((el) => {
        for (const key in el) {
          if (el[key].includes("МОЛ")) return el;
        }
      });

      //* Find header places to use them as column key
      const headerPlaces = findHeadersPlaces(data.slice(0, startIndex));

      //* Parse data
      const parsedData = parseData(data, headerPlaces);
      const groupedData = groupDataByFIO(parsedData, startIndex);

      //* Check directory for account dir exist
      const checkAccountPath = await checkDirectory(
        `parseResult/${headerPlaces.account.slice(0, 2)}`
      );
      if (!checkAccountPath)
        await fs.promises.mkdir(
          `parseResult/${headerPlaces.account.slice(0, 2)}`,
          {
            recursive: true,
          }
        );

      //* Check directory for file dir exist
      const checkPath = await checkDirectory(
        `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`
      );
      if (!checkPath)
        await fs.promises.mkdir(
          `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`,
          { recursive: true }
        );

      //* Creating new xlsx files foreach person
      for (let i = 0; i < groupedData.length; i++) {
        switch (headerPlaces.account.slice(0, 2)) {
          case "10":
            await templates.getExcelFor10(
              headerPlaces.account,
              groupedData[i].FIO,
              groupedData[i].data,
              `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`
            );
            break;
          case "11":
            await templates.getExcelFor11(
              headerPlaces.account,
              groupedData[i].FIO,
              groupedData[i].data,
              `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`
            );
            break;
          case "12":
            break;
          case "22":
            await templates.getExcelFor22(
              headerPlaces.account,
              groupedData[i].FIO,
              groupedData[i].data,
              `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`
            );
            break;
          case "23":
            await templates.getExcelFor23(
              headerPlaces.account,
              groupedData[i].FIO,
              groupedData[i].data,
              `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`
            );
            break;
          default:
            throw new Error(
              "Формулы расчеты для такого счета не существует в программе!"
            );
        }
      }
    } catch (error) {
      console.log(new Error(error.message));
    } finally {
      rl.close();
    }
  });
};

start();
