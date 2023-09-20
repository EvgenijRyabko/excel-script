export const parseData = (
  data,
  /** @type { { name: string; price: string; units: string; rest: string; account: string } } */ headerPlaces
) => {
  return data.map((el) => {
    return {
      name: el[headerPlaces.name],
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
