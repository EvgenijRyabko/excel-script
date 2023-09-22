import * as xlsx from "xlsx/xlsx.mjs";
import * as fs from "fs";

export const parseFile = async (file, monthName) => {
  const buf = fs.readFileSync(file);
  const workbook = xlsx.read(buf);
  const sheet_name_list = workbook.SheetNames;
  const monthIndex = sheet_name_list.findIndex((el) =>
    el.toLowerCase().includes(monthName)
  );

  if (monthIndex <= 0)
    throw new Error(
      `В файле отсутствует данные для данного месяца либо лист не назван именем месяца!.`
    );
  return xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[monthIndex]]);
};

export const parseData = (
  data,
  /** @type { { name: string; price: string; units: string; rest: string; account: string; nomNumber: string } } */ headerPlaces
) => {
  return data.map((el) => {
    return {
      nomNumber: `${el[headerPlaces.nomNumber] || ""}`,
      name: el[headerPlaces.name]?.trimStart(),
      units: el[headerPlaces.units],
      price: el[headerPlaces.price],
      restForMonthEnd_Amount: el[headerPlaces.rest],
      restForMonthEnd_Price:
        el[
          `${headerPlaces.rest.split("EMPTY_")[0]}EMPTY_${
            parseInt(headerPlaces.rest.split("EMPTY_")[1]) + 1
          }`
        ],
    };
  });
};

export const groupDataByFIO = (data, startIndex) => {
  /** @type { {FIO: string, data: any[]}[]} */
  const result = [];

  for (let i = startIndex; i < data.length; i++) {
    if (!data[i].name) continue;
    if (data[i].name.includes("МОЛ")) {
      result.push({
        FIO: data[i].name
          .split(":")
          .reduce((arr, el, i) => (i === 0 ? arr : (arr = [...arr, el])), [])
          .join(" ")
          .trimStart(),
        data: [],
      });
    } else {
      result[result.length - 1].data.push(data[i]);
    }
  }

  return result;
};
