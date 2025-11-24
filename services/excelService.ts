import * as XLSXPkg from 'xlsx';
import { LMSDataRow, MasterDataRow, FinalReportRow } from '../types';

// Robustly resolve the XLSX library object
const getXLSX = (): any => {
  try {
    // 1. Try global window object first (Most reliable with the script tag in index.html)
    if (typeof window !== 'undefined' && (window as any).XLSX) {
       const globalLib = (window as any).XLSX;
       if (globalLib && (globalLib.read || globalLib.utils)) return globalLib;
    }

    // 2. Try the imported package (Fallback)
    const pkg = XLSXPkg as any;
    if (pkg && pkg.read && pkg.utils) return pkg;
    if (pkg && pkg.default && pkg.default.read && pkg.default.utils) return pkg.default;
    
  } catch (e) {
    console.warn("Error resolving XLSX library:", e);
  }
  
  return null;
};

// Column Mapping Configuration
const COLUMN_ALIASES: Record<string, string> = {
  // Email Variants - IMPORTANT: Maps "Email address" (LMS) to "Username (Email)" (Master)
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

// Helper to clean text: remove hidden spaces, non-breaking spaces, trim, and lowercase
const cleanText = (val: any): string => {
  if (val === null || val === undefined) return "";
  return String(val)
    .replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, ' ') // Remove non-breaking spaces
    .replace(/\s+/g, ' ') // Collapse multiple spaces to one
    .trim()
    .toLowerCase();
};

// Helper to normalize keys: trim spaces, remove hidden characters, and apply alias mapping
const normalizeKeys = (obj: any): any => {
  const newObj: any = {};
  Object.keys(obj).forEach((key) => {
    let cleanKey = key.trim();
    
    // Apply Alias Mapping to standardize column names
    // 1. Direct Match
    if (COLUMN_ALIASES[cleanKey]) {
      cleanKey = COLUMN_ALIASES[cleanKey];
    } else {
      // 2. Case-insensitive Match (Helper for slightly different headers)
      const lowerKey = cleanKey.toLowerCase();
      const aliasMatch = Object.keys(COLUMN_ALIASES).find(k => k.toLowerCase() === lowerKey);
      if (aliasMatch) {
        cleanKey = COLUMN_ALIASES[aliasMatch];
      }
    }

    if (cleanKey) {
      const val = obj[key];
      // If multiple source columns map to the same target (e.g., legacy columns), preserve existing non-empty data
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
        
        // Ensure library is loaded
        const lib = getXLSX();
        
        if (!lib) {
           throw new Error("Excel library (XLSX) not found. Please refresh the page.");
        }
        
        if (!lib.read || !lib.utils) {
          throw new Error("XLSX library failed to initialize correctly. Please refresh the page.");
        }

        const workbook = lib.read(data, { type: 'array' });
        
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error("The uploaded file appears to be empty or invalid.");
        }

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        
        if (!sheet) {
          throw new Error("Sheet could not be loaded. Please confirm the file is not empty and the header row exists.");
        }

        // Convert to JSON with raw values first. defval: "" ensures empty cells have keys.
        const json = lib.utils.sheet_to_json(sheet, { defval: "" });
        
        // Normalize keys to handle aliases and whitespace issues
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
  talent: LMSDataRow[],
  pharmacy: LMSDataRow[],
  master: MasterDataRow[]
): FinalReportRow[] => {
  // 1. Combine Raw LMS sheets from Talent and Pharmacy
  const combinedLMS = [...talent, ...pharmacy];

  // 2. Identify Lesson Columns dynamically based on scanning cell values
  const lessonColumns = new Set<string>();
  const allKeys = new Set<string>();

  // Gather all unique keys from the entire dataset
  combinedLMS.forEach(row => Object.keys(row).forEach(k => allKeys.add(k)));

  const TARGET_VALUE_COMPLETED = "completed (achieved pass grade)";
  const TARGET_VALUE_COMPLETED_SIMPLE = "completed"; // Also accept just "Completed"
  const TARGET_VALUE_NOT_COMPLETED = "not completed";

  allKeys.forEach(key => {
    // Scan all rows for this specific column key to see if it contains lesson status values
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

  // Helper to calculate completion rate for a single row
  const calculateRate = (row: LMSDataRow): number => {
    if (lessonColumns.size === 0) return 0;
    
    let completedCount = 0;
    
    lessonColumns.forEach(colKey => {
      const rawVal = (row as any)[colKey];
      const val = cleanText(rawVal);
      
      // Count if it matches either the long form or strict "completed"
      if (val === TARGET_VALUE_COMPLETED || val === TARGET_VALUE_COMPLETED_SIMPLE) {
        completedCount++;
      }
    });

    // Return a decimal (0.0 to 1.0)
    return completedCount / lessonColumns.size; 
  };

  // 3. Create a Map for LMS data indexed by EMAIL (LMS is the primary source of truth for "Existence")
  const lmsMap = new Map<string, { row: LMSDataRow, rate: number }>();

  combinedLMS.forEach((row) => {
    const rawEmail = row['Username (Email)'];
    if (!rawEmail || typeof rawEmail !== 'string') return;
    const emailKey = cleanText(rawEmail); 
    if (!emailKey || emailKey === 'undefined') return;

    // Calculate rate
    const rate = calculateRate(row);

    const existing = lmsMap.get(emailKey);

    if (existing) {
      // Priority: Higher completion rate
      if (rate > existing.rate) {
        lmsMap.set(emailKey, { row, rate });
      } else if (rate === existing.rate) {
        // Tie-breaker: Latest date
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

  // 4. Create a Map for MASTER data for fast metadata lookup
  const masterMap = new Map<string, MasterDataRow>();
  master.forEach((row) => {
    const rawEmail = row['Username (Email)'];
    if (rawEmail) {
       masterMap.set(cleanText(rawEmail), row);
    }
  });

  const finalRows: FinalReportRow[] = [];

  // 5. Generate Final Report
  // LOGIC: Iterate through LMS (Raw Data) map. 
  // - Included: Everyone in Raw Data (Intersection + Raw Only).
  // - Excluded: People in Master Sheet who are NOT in Raw Data (Rate would be unknown/0, but user requested raw-based logic).
  
  lmsMap.forEach((lmsEntry, emailKey) => {
    const lmsRow = lmsEntry.row;
    const completionRate = lmsEntry.rate;

    // Attempt to find in Master Sheet to enrich data
    const masterRow = masterMap.get(emailKey);

    // Prefer Master data if available, fallback to LMS data
    const district = masterRow?.District || lmsRow.District || '';
    const city = masterRow?.City || lmsRow.City || '';
    const supervisor = masterRow?.['Supervisor Name'] || lmsRow['Supervisor Name'] || '';
    const pharmacyNo = masterRow?.['Pharmacy No.'] || lmsRow['Pharmacy No.'] || '';
    const employeeId = masterRow?.['User/Employee ID'] || lmsRow['User/Employee ID'] || '';
    const scfhs = masterRow?.SCFHS || lmsRow.SCFHS || '';
    
    const displayName = masterRow?.['Display Name (Pharmacist name)'] || lmsRow['Display Name (Pharmacist name)'] || '';
    const phone = masterRow?.['Phone number (Whatsapp)'] || lmsRow['Phone number (Whatsapp)'] || '';
    const email = masterRow?.['Username (Email)'] || lmsRow['Username (Email)'] || '';

    // Only add if we have some valid identifier
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

export const generateExcelFile = (data: FinalReportRow[]) => {
  const lib = getXLSX();
  
  // Safe check
  if (!lib?.utils) {
    console.error("XLSX library not loaded properly, cannot generate file.");
    alert("Error: XLSX library not available. Please refresh the page.");
    return;
  }
  
  const wb = lib.utils.book_new();

  // =================================================================================
  // SHEET 1: DASHBOARD
  // =================================================================================

  const total = data.length;
  // Use slightly lenient check for floating point precision (e.g. 0.99999)
  const completed = data.filter(d => d["Completion Rate"] >= 0.999).length; 
  const inProgress = data.filter(d => d["Completion Rate"] > 0 && d["Completion Rate"] < 0.999).length;
  const notStarted = data.filter(d => d["Completion Rate"] === 0).length;

  const summaryHeaders = ["Summarization Status", "Number Count", "Percentage"];
  const summaryData = [
    summaryHeaders,
    ["Number Of Student", total, 1], // 100%
    ["Number Of Student Who Completed", completed, total > 0 ? completed / total : 0],
    ["Number OF Student Who In Progress", inProgress, total > 0 ? inProgress / total : 0],
    ["Number Of Students Who Did Not Started", notStarted, total > 0 ? notStarted / total : 0]
  ];

  const wsDashboard = lib.utils.json_to_sheet(data, { origin: "A9" });

  lib.utils.sheet_add_aoa(wsDashboard, summaryData, { origin: "C2" });

  // Formatting
  ['E3', 'E4', 'E5', 'E6'].forEach(cellRef => {
    if (wsDashboard[cellRef]) wsDashboard[cellRef].z = '0%';
  });

  const range = lib.utils.decode_range(wsDashboard['!ref'] || "A1:A1");
  let completionColIndex = -1;
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellAddress = lib.utils.encode_cell({ r: 8, c: C });
    if (wsDashboard[cellAddress] && wsDashboard[cellAddress].v === "Completion Rate") {
      completionColIndex = C;
      break;
    }
  }
  
  if (completionColIndex !== -1) {
    for (let R = 9; R <= range.e.r; ++R) {
      const cellAddress = lib.utils.encode_cell({ r: R, c: completionColIndex });
      if (wsDashboard[cellAddress]) wsDashboard[cellAddress].z = '0%';
    }
  }

  wsDashboard['!cols'] = [
    { wch: 15 }, // A: District
    { wch: 15 }, // B: City
    { wch: 30 }, // C: Supervisor
    { wch: 20 }, // D: Pharmacy
    { wch: 20 }, // E: ID
    { wch: 25 }, // F: Email
    { wch: 25 }, // G: Name
    { wch: 15 }, // H: Phone
    { wch: 15 }, // I: SCFHS
    { wch: 15 }, // J: Completion
  ];

  wsDashboard['!autofilter'] = { 
    ref: lib.utils.encode_range({
      s: { r: 8, c: 0 }, 
      e: { r: range.e.r, c: range.e.c } 
    }) 
  };

  // =================================================================================
  // SHEET 2: PIVOT ANALYSIS
  // =================================================================================

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
        .map(([key, stats]) => [
            key, 
            stats.total, 
            stats.comp, 
            stats.prog, 
            stats.not, 
            stats.total ? stats.comp / stats.total : 0
        ]);
        
    return [
        [title, "", "", "", "", ""],
        [groupByField, "Total Count", "Completed", "In Progress", "Not Started", "Completion %"],
        ...rows,
        [""] 
    ];
  };

  const pivotDistrict = createPivotTable("District", "Pivot: Breakdown by District");
  const pivotSupervisor = createPivotTable("Supervisor Name", "Pivot: Breakdown by Supervisor");
  const pivotCity = createPivotTable("City", "Pivot: Breakdown by City");

  const wsPivot = lib.utils.aoa_to_sheet([["Pivot Analysis Report"], [""]]);
  
  let currentRow = 2;
  [pivotDistrict, pivotSupervisor, pivotCity].forEach(tableData => {
      lib.utils.sheet_add_aoa(wsPivot, tableData, { origin: { r: currentRow, c: 0 } });
      
      const rowsInTable = tableData.length;
      for (let i = 2; i < rowsInTable - 1; i++) {
         const r = currentRow + i;
         const ref = lib.utils.encode_cell({ r, c: 5 });
         if (!wsPivot[ref]) wsPivot[ref] = { v: 0, t: 'n' };
         wsPivot[ref].z = '0%';
      }

      currentRow += tableData.length + 1;
  });

  wsPivot['!cols'] = [{ wch: 30 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 }];
  
  const pivotRange = lib.utils.decode_range(wsPivot['!ref'] || "A1:F100");
  wsPivot['!autofilter'] = { 
    ref: lib.utils.encode_range(pivotRange) 
  };

  lib.utils.book_append_sheet(wb, wsDashboard, "Dashboard Report");
  lib.utils.book_append_sheet(wb, wsPivot, "Pivot Analysis");

  lib.writeFile(wb, "United_Pharmacy_Completion_Report.xlsx");
};