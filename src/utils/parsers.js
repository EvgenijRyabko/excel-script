export const parseData = (data) => {
  return data.map((el) => {
    console.log(el);
    delete el["__EMPTY"];

    let name;

    for (const key in el) {
      if (key.includes("ОБОРОТНАЯ")) name = el[key];
    }

    return {
      name: name || "",
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

export const groupDataByFIO = (data, startIndex) => {
  /** @type { {FIO: string, data: any[]}[]} */
  const result = [];

  for (let i = startIndex; i < data.length; i++) {
    if (data[i].name.includes("МОЛ")) {
      result.push({
        FIO: data[i].name.split(":")[1].trimStart(),
        data: [],
      });
    } else {
      result[result.length - 1].data.push(data[i]);
    }
  }

  return result;
};
