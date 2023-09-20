import { checkDirectory } from "./utils/checkDir.js";
import { rl } from "./utils/readline.js";
import { writeFile } from "./utils/writeFile.js";
import { parseData, groupDataByFIO } from "./utils/parsers.js";
import { findHeadersPlaces } from "./utils/findHeadersPlaces.js";
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

      //* Check directory for file exist
      const checkPath = await checkDirectory(`parseResult/${fileName}`);
      if (!checkPath)
        await fs.promises.mkdir(`parseResult/${fileName}`, { recursive: true });

      //* Creating new xlsx files foreach person
      for (let i = 0; i < groupedData.length; i++) {
        await writeFile(
          headerPlaces.account,
          groupedData[i].FIO,
          groupedData[i].data,
          `parseResult/${fileName}`
        );
      }
    } catch (error) {
      console.log(new Error(error.message));
    } finally {
      rl.close();
    }
  });
};

start();
