export const findHeadersPlaces = (headers) => {
  let name, price, units, rest, account, nomNumber;

  for (let i = 0; i < headers.length; i++) {
    for (const key in headers[i]) {
      if (headers[i][key].includes("по счету"))
        account = headers[i][key].split(" ")[2];
      if (headers[i][key].includes("Название")) name = key;
      if (headers[i][key].includes("Цена в руб.")) price = key;
      if (["Ед. изм.", "Ед.изм"].some((el) => headers[i][key].includes(el)))
        units = key;
      if (headers[i][key].includes("Остаток на конец месяца")) rest = key;
      if (
        [
          "Инвен.номер",
          "Инвентарный номер",
          "Артикул",
          "Номенклатурный номер",
        ].some((el) => headers[i][key].includes(el))
      )
        nomNumber = key;
    }
  }

  return { name, price, units, rest, account, nomNumber };
};
