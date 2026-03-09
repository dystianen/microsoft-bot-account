const XLSX = require("xlsx");

const EXCEL_FILE = "./accounts.xlsx";

let writeLock = Promise.resolve();

function writeResultToExcel(rowIndex, status, domainEmail) {
  writeLock = writeLock.then(() => {
    try {
      const workbook = XLSX.readFile(EXCEL_FILE);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Find column positions from existing headers (row 1)
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      let statusCol = -1;
      let domainCol = -1;

      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 0, c });
        const cell = sheet[cellRef];
        if (!cell) continue;
        const val = String(cell.v).trim().toLowerCase();
        if (val === "status") statusCol = c;
        if (val === "domain email") domainCol = c;
      }

      if (statusCol === -1 || domainCol === -1) {
        console.error(
          '[Excel] Could not find "Status" or "Domain Email" column in header row.',
        );
        console.error(
          'Please add "Status" and "Domain Email" columns to accounts.xlsx first!',
        );
        return;
      }

      console.log(
        `[Excel] Found columns -> Status: col ${statusCol}, Domain Email: col ${domainCol}`,
      );

      // Write only the Status and Domain Email cells
      const dataRow = rowIndex + 1; // 0-indexed, row 0 = header
      const statusCellRef = XLSX.utils.encode_cell({
        r: dataRow,
        c: statusCol,
      });
      const domainCellRef = XLSX.utils.encode_cell({
        r: dataRow,
        c: domainCol,
      });

      sheet[statusCellRef] = { t: "s", v: status };
      sheet[domainCellRef] = { t: "s", v: domainEmail || "" };

      if (!sheet["!cols"]) {
        sheet["!cols"] = [
          { wch: 35 },
          { wch: 35 },
          { wch: 15 },
          { wch: 15 },
          { wch: 30 },
          { wch: 15 },
          { wch: 15 },
          { wch: 25 },
          { wch: 30 },
          { wch: 15 },
          { wch: 15 },
          { wch: 12 },
          { wch: 15 },
          { wch: 20 },
          { wch: 8 },
          { wch: 10 },
          { wch: 10 },
          { wch: 12 },
          { wch: 40 },
        ];
      }

      ws["!cols"] = [
        { wch: 35 }, // Email
        { wch: 35 }, // Password
        { wch: 15 }, // First Name
        { wch: 15 }, // Last Name
        { wch: 30 }, // Company Name
        { wch: 15 }, // Company Size
        { wch: 15 }, // Phone
        { wch: 25 }, // Job Title
        { wch: 30 }, // Address
        { wch: 15 }, // City
        { wch: 15 }, // State
        { wch: 12 }, // Postal Code
        { wch: 15 }, // Country
        { wch: 20 }, // Card Number
        { wch: 8 }, // CVV
        { wch: 10 }, // Exp Month
        { wch: 10 }, // Exp Year
        { wch: 12 }, // Status
        { wch: 40 }, // Domain Email
      ];
      XLSX.writeFile(workbook, EXCEL_FILE);

      const excelRow = rowIndex + 2;
      console.log(
        `[Excel] Row ${excelRow} updated: Status=${status}, Domain=${domainEmail || "N/A"}`,
      );
    } catch (err) {
      console.error(`[Excel] Failed:`, err.message);
    }
  });
  return writeLock;
}

async function test() {
  console.log("=== Testing writeResultToExcel ===\n");

  // Row 0 = first data row -> SUCCESS with domain email
  await writeResultToExcel(
    0,
    "SUCCESS",
    "orion@detroitautomotive.onmicrosoft.com",
  );

  // Row 1 = second data row -> FAILED
  await writeResultToExcel(1, "FAILED", "");

  console.log("\n✅ Test done! Open accounts.xlsx to verify.");
}

test();
