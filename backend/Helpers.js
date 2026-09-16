/* ============================================================
   HELPER
   ============================================================ */

function deleteRowsByKey_(sheet, keyColumn, keyValue) {
  const lastRow = sheet.getLastRow();
  if (lastRow<=1) return;
  const rows = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn()).getValues();
  for (let i=rows.length-1;i>=0;i--) {
    if (String(rows[i][keyColumn-1])===String(keyValue)) sheet.deleteRow(i+2);
  }
}

function getSheet_(
  ss,
  sheetName
) {

  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    throw new Error(
      "Sheet '" +
      sheetName +
      "' tidak ditemukan."
    );

  }


  return sheet;

}

function getDataRows_(
  sheet
) {

  const lastRow =
    sheet.getLastRow();


  const lastColumn =
    sheet.getLastColumn();


  if (
    lastRow <= 1 ||
    lastColumn <= 0
  ) {

    return [];

  }


  return sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      lastColumn
    )
    .getValues();

}

function getNextSequenceStart_(
  sheet,
  prefix
) {

  const rows =
    getDataRows_(
      sheet
    );

  let maxNumber =
    0;

  rows.forEach(function(row) {

    const value =
      cleanString_(row[0]);

    const match =
      value.match(
        new RegExp("^" + prefix + "(\\d+)$", "i")
      );

    if (match) {
      const number = Number(match[1]);
      if (number > maxNumber) {
        maxNumber = number;
      }
    }

  });

  const props =
    PropertiesService
      .getScriptProperties();

  const key =
    "ID_COUNTER_" + prefix;

  const propertyCounter =
    Number(
      props.getProperty(key)
    ) || 0;

  return Math.max(
    maxNumber,
    propertyCounter
  ) + 1;

}

function formatSequentialId_(
  prefix,
  number
) {

  return (
    prefix +
    String(number)
      .padStart(4, "0")
  );

}

function getNextId_(
  sheet,
  prefix
) {

  const rows =
    getDataRows_(
      sheet
    );


  let maxNumber =
    0;


  rows.forEach(
    function(row) {

      const value =
        cleanString_(
          row[0]
        );


      const match =
        value.match(
          new RegExp(
            "^" +
            prefix +
            "(\\d+)$",
            "i"
          )
        );


      if (match) {

        const number =
          Number(
            match[1]
          );


        if (
          number >
          maxNumber
        ) {

          maxNumber =
            number;

        }

      }

    }
  );


  /*
   * Simpan juga counter di Script Properties
   * agar tetap cepat dan aman dari duplikasi.
   */

  const props =
    PropertiesService
      .getScriptProperties();


  const key =
    "ID_COUNTER_" +
    prefix;


  const propertyCounter =
    Number(
      props.getProperty(key)
    ) || 0;


  const nextNumber =
    Math.max(
      maxNumber,
      propertyCounter
    ) + 1;


  props.setProperty(
    key,
    String(nextNumber)
  );


  return (
    prefix +
    String(nextNumber)
      .padStart(4, "0")
  );

}

function normalizeKategoriPenjualan_(
  value
) {

  const text =
    cleanString_(
      value
    )
    .toLowerCase();


  if (
    text === "grey"
  ) {

    return "Grey";

  }


  if (
    text === "kain jadi"
  ) {

    return "Kain Jadi";

  }


  return "";

}

function cleanString_(
  value
) {

  if (
    value === null ||
    value === undefined
  ) {

    return "";

  }


  return String(
    value
  ).trim();

}

function toNumber_(
  value
) {

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {

    return 0;

  }


  if (
    typeof value ===
    "number"
  ) {

    return isNaN(value)
      ? 0
      : value;

  }


  let text =
    String(value)
      .trim();


  text =
    text
      .replace(
        /Rp/gi,
        ""
      )
      .replace(
        /\s/g,
        ""
      );


  if (
    text.includes(".") &&
    text.includes(",")
  ) {

    text =
      text
        .replace(
          /\./g,
          ""
        )
        .replace(
          ",",
          "."
        );

  } else if (
    text.includes(".")
  ) {

    const parts =
      text.split(".");


    if (
      parts.length === 2 &&
      parts[1].length === 3
    ) {

      text =
        text.replace(
          /\./g,
          ""
        );

    }

  } else if (
    text.includes(",")
  ) {

    text =
      text.replace(
        ",",
        "."
      );

  }


  const number =
    Number(text);


  return isNaN(number)
    ? 0
    : number;

}

function parseDate_(
  value
) {

  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {

    return value;

  }


  if (!value) {

    throw new Error(
      "Tanggal tidak valid."
    );

  }


  const text =
    String(value)
      .trim();


  const iso =
    text.match(
      /^(\d{4})-(\d{2})-(\d{2})$/
    );


  if (iso) {

    return new Date(
      Number(iso[1]),
      Number(iso[2]) - 1,
      Number(iso[3])
    );

  }


  const indo =
    text.match(
      /^(\d{2})\/(\d{2})\/(\d{4})$/
    );


  if (indo) {

    return new Date(
      Number(indo[3]),
      Number(indo[2]) - 1,
      Number(indo[1])
    );

  }


  const date =
    new Date(text);


  if (
    isNaN(
      date.getTime()
    )
  ) {

    throw new Error(
      "Format tanggal tidak valid: " +
      text
    );

  }


  return date;

}

function formatDateForClient_(
  value
) {

  if (!value) {

    return "";

  }


  let date;


  if (
    value instanceof Date
  ) {

    date =
      value;

  } else {

    date =
      new Date(value);

  }


  if (
    isNaN(
      date.getTime()
    )
  ) {

    return "";

  }


  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

}

function parseClientDate_(
  value
) {

  if (!value) {

    return null;

  }


  try {

    return parseDate_(
      value
    );

  } catch (error) {

    return null;

  }

}

function isSameMonth_(
  dateValue,
  year,
  month
) {

  const date =
    parseClientDate_(
      dateValue
    );


  if (!date) {

    return false;

  }


  return (
    date.getFullYear() ===
    Number(year)
  ) &&
  (
    date.getMonth() + 1 ===
    Number(month)
  );

}

function formatDateTimeForClient_(value) {
  if (!value) return "";
  let date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function monthName_(
  month
) {

  const names = [

    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember"

  ];


  return (
    names[
      Number(month) - 1
    ] || ""
  );

}


