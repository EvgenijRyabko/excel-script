import { checkDirectory } from "./utils/checkDir.js";
import { parseData, groupDataByFIO } from "./utils/parsers.js";
import { findHeadersPlaces } from "./utils/findHeadersPlaces.js";
import { parseFile } from "./utils/parsers.js";
import templates from "./templates/index.js";
import * as fs from "fs";

const start = async () => {
  //* Переменные для статистики
  let fileCount = 0;
  let skippedFiles = 0;

  fs.readdir("filesToParse", (err, files) => {
    try {
      if (err) throw new Error(err);

      files.forEach(async (file) => {
        try {
          if (!file.includes("xls")) {
            throw new Error(
              `[${file}]: Файл с таким типом не подходит скрипту!`
            );
          }

          const fileName = file.split(".xls")[0];

          //* Get data from excel file
          const data = await parseFile(`filesToParse/${fileName}.xlsx`, "июнь");

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
            const checkPath = await checkDirectory(
              `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}/${
                groupedData[i].FIO
              }.xlsx`
            );

            if (checkPath) {
              skippedFiles = skippedFiles + 1;
              continue;
            }

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
                console.log(
                  "\n-\t\t\t-\nДля 12 счета case еще не написан\n-\t\t\t-\n"
                );
                break;
              case "20":
                await templates.getExcelFor20(
                  headerPlaces.account,
                  groupedData[i].FIO,
                  groupedData[i].data,
                  `parseResult/${headerPlaces.account.slice(0, 2)}/${fileName}`
                );
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
                  "\x1b[31mФормулы расчеты для такого счета не существует в программе!" +
                    `\nФайл: ${`parseResult/${headerPlaces.account.slice(
                      0,
                      2
                    )}/${fileName}/${groupedData[i].FIO}\x1b[0m`}`
                );
            }

            fileCount = fileCount + 1;
          }

          console.log(`\n\tФайл ${fileName}`);
          console.log("\x1b[32m", `Обработано ${fileCount} файлов\x1b[0m`);
          console.log("\x1b[38;5;33m", `Пропущено: ${skippedFiles}\x1b[0m`);
          fileCount = 0;
          skippedFiles = 0;
        } catch (e) {
          console.log("\x1b[31m", ` File [${file}]: ${e.message}\x1b[0m`);
        }
      });
    } catch (error) {
      console.log("\x1b[31m", `${error.message}\x1b[0m`);
    }
  });
};

start();
