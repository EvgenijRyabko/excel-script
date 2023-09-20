export const findHeadersPlaces = (headers) => {
  let name, price, units, rest, account;

  for (let i = 0; i < headers.length; i++) {
    for (const key in headers[i]) {
      if (headers[i][key].includes("по счету"))
        account = headers[i][key].split(" ")[2];
      if (headers[i][key].includes("Название")) name = key;
      if (headers[i][key].includes("Цена в руб.")) price = key;
      if (headers[i][key].includes("Ед. изм.")) units = key;
      if (headers[i][key].includes("Остаток на конец месяца")) rest = key;
    }
  }

  return { name, price, units, rest, account };
};
