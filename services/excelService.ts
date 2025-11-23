import * as XLSXPkg from 'xlsx';
import { LMSDataRow, MasterDataRow, FinalReportRow } from '../types';

// Robustly resolve the XLSX library object
const getXLSX = (): any => {
  try {
    // 1. Try global window object first (Most reliable for xlsx-js-style via script tag)
    // This is preferred because 'xlsx-js-style' modifies the global object to add style support.
    if (typeof window !== 'undefined' && (window as any).XLSX) {
       const globalLib = (window as any).XLSX;
       if (globalLib.utils) return globalLib;
    }

    // 2. Try the imported package
    // When importing a UMD bundle as a module, the namespace object might vary.
    const pkg = XLSXPkg as any;
    
    if (pkg) {
        // Check if 'utils' exists directly
        if (pkg.utils) return pkg;
        
        // Check default export
        if (pkg.default && pkg.default.utils) return pkg.default;

        // Check named export 'XLSX'
        if (pkg.XLSX && pkg.XLSX.utils) return pkg.XLSX;
    }

    return null;
  } catch (e) {
    console.warn("Error resolving XLSX library:", e);
    return null;
  }
};

// Column Mapping Configuration
const COLUMN_ALIASES: Record<string, string> = {
  // Email Variants
  "Email address": "Username (Email)",
  "Username": "Username (Email)",
  "Email": "Username (Email)",
  "User Email": "Username (Email)",
  "E-mail": "Username (Email)",
  
  // ID Variants
  "User ID": "User/Employee ID",
  "Employee ID": "User/Employee ID",
  "EmployeeID": "User/Employee ID",
  "UserID": "User/Employee ID",
  "User/EmployeeID": "User/Employee ID",
  "ID": "User/Employee ID",
  "Emp ID": "User/Employee ID",
  
  // Name Variants
  "Display Name": "Display Name (Pharmacist name)",
  "Pharmacist Name": "Display Name (Pharmacist name)",
  "Pharmacist": "Display Name (Pharmacist name)",
  "Full Name": "Display Name (Pharmacist name)",
  "Name": "Display Name (Pharmacist name)",
  
  // Phone Variants
  "Phone": "Phone number (Whatsapp)",
  "Mobile": "Phone number (Whatsapp)",
  "Phone Number": "Phone number (Whatsapp)",
  "Whatsapp": "Phone number (Whatsapp)",
  "Contact No": "Phone number (Whatsapp)",

  // Pharmacy No Variants
  "Pharmacy ID": "Pharmacy No.",
  "Pharmacy #": "Pharmacy No.",
  "Pharmacy Code": "Pharmacy No.",

  // Supervisor Variants
  "Supervisor": "Supervisor Name",
  "Manager": "Supervisor Name"
};

// Helper to clean text
const cleanText = (val: any): string => {
  if (val === null || val === undefined) return "";
  return String(val)
    .replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, ' ') 
    .replace(/\s+/g, ' ') 
    .trim()
    .toLowerCase();
};

// Helper to normalize keys
const normalizeKeys = (obj: any): any => {
  const newObj: any = {};
  Object.keys(obj).forEach((key) => {
    let cleanKey = key.trim();
    
    if (COLUMN_ALIASES[cleanKey]) {
      cleanKey = COLUMN_ALIASES[cleanKey];
    } else {
      const lowerKey = cleanKey.toLowerCase();
      const aliasMatch = Object.keys(COLUMN_ALIASES).find(k => k.toLowerCase() === lowerKey);
      if (aliasMatch) {
        cleanKey = COLUMN_ALIASES[aliasMatch];
      }
    }

    if (cleanKey) {
      const val = obj[key];
      if (newObj[cleanKey] === undefined || (newObj[cleanKey] === "" && val !== "" && val !== undefined)) {
        newObj[cleanKey] = val;
      }
    }
  });
  return newObj;
};

export const parseExcel = async <T>(file: File): Promise<T[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const lib = getXLSX();
        
        if (!lib || !lib.utils) {
           throw new Error("Excel library (XLSX) not found. Please ensure your internet connection allows loading external scripts.");
        }

        const workbook = lib.read(data, { type: 'array' });
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error("The uploaded file appears to be empty or invalid.");
        }

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        
        if (!sheet) {
          throw new Error("Sheet could not be loaded.");
        }

        const json = lib.utils.sheet_to_json(sheet, { defval: "" });
        const normalized = json.map((row: any) => normalizeKeys(row)) as T[];
        resolve(normalized);
      } catch (error) {
        console.error("Parse Error:", error);
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

export const processData = (
  lms1: LMSDataRow[],
  lms2: LMSDataRow[],
  master: MasterDataRow[]
): FinalReportRow[] => {
  // 1. Combine Raw LMS sheets
  const combinedLMS = [...lms1, ...lms2];

  // 2. Identify Lesson Columns
  const lessonColumns = new Set<string>();
  const allKeys = new Set<string>();

  combinedLMS.forEach(row => Object.keys(row).forEach(k => allKeys.add(k)));

  const TARGET_VALUE_COMPLETED = "completed (achieved pass grade)";
  const TARGET_VALUE_COMPLETED_SIMPLE = "completed";
  const TARGET_VALUE_NOT_COMPLETED = "not completed";

  allKeys.forEach(key => {
    const isLessonCol = combinedLMS.some(row => {
      const rawVal = (row as any)[key];
      const val = cleanText(rawVal);
      return val === TARGET_VALUE_COMPLETED || 
             val === TARGET_VALUE_NOT_COMPLETED || 
             val === TARGET_VALUE_COMPLETED_SIMPLE;
    });

    if (isLessonCol) {
      lessonColumns.add(key);
    }
  });

  const calculateRate = (row: LMSDataRow): number => {
    if (lessonColumns.size === 0) return 0;
    let completedCount = 0;
    lessonColumns.forEach(colKey => {
      const rawVal = (row as any)[colKey];
      const val = cleanText(rawVal);
      if (val === TARGET_VALUE_COMPLETED || val === TARGET_VALUE_COMPLETED_SIMPLE) {
        completedCount++;
      }
    });
    return completedCount / lessonColumns.size; 
  };

  // 3. Create Map for LMS data
  const lmsMap = new Map<string, { row: LMSDataRow, rate: number }>();

  combinedLMS.forEach((row) => {
    const rawEmail = row['Username (Email)'];
    if (!rawEmail || typeof rawEmail !== 'string') return;
    const emailKey = cleanText(rawEmail); 
    if (!emailKey || emailKey === 'undefined') return;

    const rate = calculateRate(row);
    const existing = lmsMap.get(emailKey);

    if (existing) {
      if (rate > existing.rate) {
        lmsMap.set(emailKey, { row, rate });
      } else if (rate === existing.rate) {
        const dateA = new Date(row.Date).getTime() || Number(row.Date) || 0;
        const dateB = new Date(existing.row.Date).getTime() || Number(existing.row.Date) || 0;
        if (dateA > dateB) {
          lmsMap.set(emailKey, { row, rate });
        }
      }
    } else {
      lmsMap.set(emailKey, { row, rate });
    }
  });

  // 4. Master Map
  const masterMap = new Map<string, MasterDataRow>();
  master.forEach((row) => {
    const rawEmail = row['Username (Email)'];
    if (rawEmail) masterMap.set(cleanText(rawEmail), row);
  });

  const finalRows: FinalReportRow[] = [];

  // 5. Generate Final Report
  lmsMap.forEach((lmsEntry, emailKey) => {
    const lmsRow = lmsEntry.row;
    const completionRate = lmsEntry.rate;

    const masterRow = masterMap.get(emailKey);

    const district = masterRow?.District || lmsRow.District || '';
    const city = masterRow?.City || lmsRow.City || '';
    const supervisor = masterRow?.['Supervisor Name'] || lmsRow['Supervisor Name'] || '';
    const pharmacyNo = masterRow?.['Pharmacy No.'] || lmsRow['Pharmacy No.'] || '';
    const employeeId = masterRow?.['User/Employee ID'] || lmsRow['User/Employee ID'] || '';
    const scfhs = masterRow?.SCFHS || lmsRow.SCFHS || '';
    
    const displayName = masterRow?.['Display Name (Pharmacist name)'] || lmsRow['Display Name (Pharmacist name)'] || '';
    const phone = masterRow?.['Phone number (Whatsapp)'] || lmsRow['Phone number (Whatsapp)'] || '';
    const email = masterRow?.['Username (Email)'] || lmsRow['Username (Email)'] || '';

    if (employeeId || email || displayName) {
      finalRows.push({
        District: String(district),
        City: String(city),
        "Supervisor Name": String(supervisor),
        "Pharmacy No.": String(pharmacyNo),
        "User/Employee ID": String(employeeId),
        "Username (Email)": String(email),
        "Display Name (Pharmacist name)": String(displayName),
        "Phone number (Whatsapp)": String(phone),
        SCFHS: String(scfhs),
        "Completion Rate": completionRate,
      });
    }
  });

  return finalRows;
};

// --- STYLING ---
const STYLES = {
  header: {
    fill: { fgColor: { rgb: "F4A460" } },
    font: { color: { rgb: "FFFFFF" }, bold: true, sz: 12, name: "Calibri" },
    alignment: { horizontal: "center", vertical: "center" },
    border: {
      top: { style: "thin", color: { rgb: "000000" } },
      bottom: { style: "thin", color: { rgb: "000000" } },
      left: { style: "thin", color: { rgb: "000000" } },
      right: { style: "thin", color: { rgb: "000000" } }
    }
  },
  cellBorder: {
    border: {
      top: { style: "thin", color: { rgb: "CCCCCC" } },
      bottom: { style: "thin", color: { rgb: "CCCCCC" } },
      left: { style: "thin", color: { rgb: "CCCCCC" } },
      right: { style: "thin", color: { rgb: "CCCCCC" } }
    },
    alignment: { vertical: "center" }
  },
  percent: {
    numFmt: "0%",
    alignment: { horizontal: "center", vertical: "center" },
    border: {
      top: { style: "thin", color: { rgb: "CCCCCC" } },
      bottom: { style: "thin", color: { rgb: "CCCCCC" } },
      left: { style: "thin", color: { rgb: "CCCCCC" } },
      right: { style: "thin", color: { rgb: "CCCCCC" } }
    }
  }
};

export const generateExcelFile = (data: FinalReportRow[]) => {
  const lib = getXLSX();
  
  if (!lib || !lib.utils) {
    alert("Error: XLSX library not available. Please refresh the page.");
    return;
  }
  
  const wb = lib.utils.book_new();

  // --- DASHBOARD SHEET ---
  const total = data.length;
  const completed = data.filter(d => d["Completion Rate"] >= 0.999).length; 
  const inProgress = data.filter(d => d["Completion Rate"] > 0 && d["Completion Rate"] < 0.999).length;
  const notStarted = data.filter(d => d["Completion Rate"] === 0).length;

  const summaryData = [
    ["Summarization Status", "Number Count", "Percentage"],
    ["Number Of Student", total, 1],
    ["Number Of Student Who Completed", completed, total > 0 ? completed / total : 0],
    ["Number OF Student Who In Progress", inProgress, total > 0 ? inProgress / total : 0],
    ["Number Of Students Who Did Not Started", notStarted, total > 0 ? notStarted / total : 0]
  ];

  const wsDashboard = lib.utils.json_to_sheet(data, { origin: "A9" });
  lib.utils.sheet_add_aoa(wsDashboard, summaryData, { origin: "C2" });

  // Style Summary Table
  const summaryRange = { s: { r: 1, c: 2 }, e: { r: 5, c: 4 } };
  for (let R = summaryRange.s.r; R <= summaryRange.e.r; ++R) {
    for (let C = summaryRange.s.c; C <= summaryRange.e.c; ++C) {
      const cellRef = lib.utils.encode_cell({ r: R, c: C });
      if (!wsDashboard[cellRef]) wsDashboard[cellRef] = { v: "" };
      
      if (R === summaryRange.s.r) {
        wsDashboard[cellRef].s = STYLES.header;
      } else {
        wsDashboard[cellRef].s = STYLES.cellBorder;
        if (C === 4) {
           wsDashboard[cellRef].z = '0%';
           wsDashboard[cellRef].s = STYLES.percent;
        }
      }
    }
  }

  // Style Main Table
  const range = lib.utils.decode_range(wsDashboard['!ref'] || "A1:A1");
  const mainHeaderRowIndex = 8;
  
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellRef = lib.utils.encode_cell({ r: mainHeaderRowIndex, c: C });
    if (wsDashboard[cellRef]) {
      wsDashboard[cellRef].s = STYLES.header;
    }
  }

  let completionColIndex = -1;
  for (let C = range.s.c; C <= range.e.c; ++C) {
     const ref = lib.utils.encode_cell({ r: mainHeaderRowIndex, c: C });
     if (wsDashboard[ref] && wsDashboard[ref].v === "Completion Rate") {
       completionColIndex = C;
     }
  }

  for (let R = mainHeaderRowIndex + 1; R <= range.e.r; ++R) {
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const cellRef = lib.utils.encode_cell({ r: R, c: C });
      if (wsDashboard[cellRef]) {
        wsDashboard[cellRef].s = STYLES.cellBorder;
        
        if (C === completionColIndex) {
          wsDashboard[cellRef].z = '0%';
          const val = wsDashboard[cellRef].v;
          let color = "000000";
          if (val >= 0.999) color = "008000";
          else if (val > 0) color = "F4A460";
          else color = "808080";

          wsDashboard[cellRef].s = {
            ...STYLES.percent,
            font: { color: { rgb: color }, bold: true }
          };
        }
      }
    }
  }

  wsDashboard['!cols'] = [
    { wch: 15 }, { wch: 15 }, { wch: 30 }, { wch: 20 }, { wch: 20 }, 
    { wch: 25 }, { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 15 }
  ];

  wsDashboard['!autofilter'] = { 
    ref: lib.utils.encode_range({ s: { r: 8, c: 0 }, e: { r: range.e.r, c: range.e.c } }) 
  };

  // --- PIVOT SHEET ---
  const createPivotTable = (groupByField: keyof FinalReportRow, title: string) => {
    const groups: Record<string, { total: number, comp: number, prog: number, not: number }> = {};
    data.forEach(row => {
        const key = String(row[groupByField] || '(Blank)');
        if (!groups[key]) groups[key] = { total: 0, comp: 0, prog: 0, not: 0 };
        groups[key].total++;
        if (row["Completion Rate"] >= 0.999) groups[key].comp++;
        else if (row["Completion Rate"] > 0 && row["Completion Rate"] < 0.999) groups[key].prog++;
        else groups[key].not++;
    });
    const rows = Object.entries(groups)
        .sort((a, b) => b[1].total - a[1].total)
        .map(([key, stats]) => [key, stats.total, stats.comp, stats.prog, stats.not, stats.total ? stats.comp / stats.total : 0]);
    return [[title, "", "", "", "", ""], [groupByField, "Total Count", "Completed", "In Progress", "Not Started", "Completion %"], ...rows, [""]];
  };

  const pivotDistrict = createPivotTable("District", "Pivot: Breakdown by District");
  const pivotSupervisor = createPivotTable("Supervisor Name", "Pivot: Breakdown by Supervisor");
  const pivotCity = createPivotTable("City", "Pivot: Breakdown by City");

  const wsPivot = lib.utils.aoa_to_sheet([["Pivot Analysis Report"], [""]]);
  let currentRow = 2;
  
  [pivotDistrict, pivotSupervisor, pivotCity].forEach(tableData => {
      lib.utils.sheet_add_aoa(wsPivot, tableData, { origin: { r: currentRow, c: 0 } });
      const rowsInTable = tableData.length;
      for (let i = 0; i < rowsInTable; i++) {
         const r = currentRow + i;
         const isHeader = (i === 1); 
         const isTitle = (i === 0);
         for(let c = 0; c < 6; c++) {
            const ref = lib.utils.encode_cell({ r, c });
            if (!wsPivot[ref]) continue;
            if (isTitle) wsPivot[ref].s = { font: { bold: true, sz: 14, color: { rgb: "F4A460" } } };
            else if (isHeader) wsPivot[ref].s = STYLES.header;
            else {
               wsPivot[ref].s = STYLES.cellBorder;
               if (c === 5 && !isHeader) {
                  wsPivot[ref].z = '0%';
                  wsPivot[ref].s = STYLES.percent;
               }
            }
         }
      }
      currentRow += tableData.length + 1;
  });

  wsPivot['!cols'] = [{ wch: 30 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }];

  lib.utils.book_append_sheet(wb, wsDashboard, "Dashboard Report");
  lib.utils.book_append_sheet(wb, wsPivot, "Pivot Analysis");

  lib.writeFile(wb, "United_Pharmacy_Completion_Report.xlsx");
};