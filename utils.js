export function ExcelDateToJSDate(date) {
  // Check whether date is a string like 12.02.2023
  // If true then return it unmanaged
  let reg = new RegExp(/^\d{2}([./-])\d{2}\1\d{4}$/);

  try {
    if (reg.test(date) === true) {
      return date;
    } else {
      // If date is in an Excel format like 75884
      // apply the formula
      let x = new Date(Math.round((date - 25569) * 86400 * 1000));
      return x.toLocaleDateString("ru-RU");
    }
  } catch {
    throw new Error("======== INVALID DATE ========");
  }
}

export function ExcelDateToJSDateLong(date) {
  let reg = new RegExp(/^\d{2}([./-])\d{2}\1\d{4}$/);
  if (reg.test(date) === true) {
    function formatDate(inputDate) {
      const months = [
        "Января",
        "Февраля",
        "Марта",
        "Апреля",
        "Мая",
        "Июня",
        "Июля",
        "Августа",
        "Сентября",
        "Октября",
        "Ноября",
        "Декабря",
      ];

      const [day, monthNum, year] = inputDate.split(".").map((str) => parseInt(str, 10));

      if (!isNaN(day) && !isNaN(monthNum) && !isNaN(year) && monthNum >= 1 && monthNum <= 12) {
        const monthName = months[monthNum - 1];
        return `${day} ${monthName} ${year}`;
      } else {
        throw new Error("Invalid date format. Please use the format: dd.mm.yyyy");
      }
    }

    return formatDate(date);
  } else {
    // If date is in an Excel format like 75884
    // apply the formula
    let x = new Date(Math.round((date - 25569) * 86400 * 1000));
    return x.toLocaleDateString("ru-RU", { dateStyle: "long" });
  }
}
