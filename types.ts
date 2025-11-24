export interface LMSDataRow {
  District: string;
  City: string;
  "Supervisor Name": string;
  Date: string | number; // Excel date might be number
  "Pharmacy No.": string;
  "User/Employee ID": string | number;
  "Username (Email)"?: string;
  "Email address"?: string;
  "Display Name (Pharmacist name)": string;
  "Phone number (Whatsapp)": string;
  SCFHS: string;
  "Attendance Status": string;
  Notes?: string;
}

export interface MasterDataRow {
  District: string;
  City: string;
  "Supervisor Name": string;
  "Pharmacy No.": string;
  "User/Employee ID": string | number;
  "Username (Email)": string;
  "Display Name (Pharmacist name)": string;
  "Phone number (Whatsapp)": string;
  SCFHS: string;
  "Completion Rate"?: number;
}

export interface FinalReportRow {
  District: string;
  City: string;
  "Supervisor Name": string;
  "Pharmacy No.": string;
  "User/Employee ID": string;
  "Username (Email)": string;
  "Display Name (Pharmacist name)": string;
  "Phone number (Whatsapp)": string;
  SCFHS: string;
  "Completion Rate": number;
}

export enum FileType {
  TALENT = 'Talent Platform Data',
  PHARMACY = 'Pharmacy Platform Data',
  MASTER = 'Master Saudi Data Sheet',
}

export interface ProcessedStats {
  totalPharmacists: number;
  totalCompleted: number;
  completionPercentage: number;
  byDistrict: { name: string; Completed: number; Pending: number }[];
  bySupervisor: { name: string; value: number }[];
  byStatus: { name: string; value: number }[];
  totalInProgress: number;
  totalNotStarted: number;
}