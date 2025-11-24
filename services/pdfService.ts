import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import { ProcessedStats, FinalReportRow } from '../types';

export const generatePDF = (stats: ProcessedStats, data: FinalReportRow[]) => {
  const doc = new jsPDF();
  const pageWidth = doc.internal.pageSize.width;

  // Title
  doc.setFontSize(22);
  doc.setTextColor(244, 164, 96); // #F4A460 Orange
  doc.text("United Pharmacy - LMS Executive Summary", 14, 20);
  
  doc.setFontSize(10);
  doc.setTextColor(100);
  doc.text(`Generated on: ${new Date().toLocaleDateString()}`, 14, 26);

  // High Level Stats
  let yPos = 40;
  doc.setFontSize(14);
  doc.setTextColor(0);
  doc.text("High-Level Overview", 14, yPos);
  
  yPos += 10;
  
  const headers = [["Metric", "Value", "Percentage"]];
  const statsData = [
    ["Total Students", stats.totalPharmacists, "100%"],
    ["Completed", stats.totalCompleted, `${Math.round(stats.completionPercentage)}%`],
    ["In Progress", stats.totalInProgress, `${Math.round((stats.totalInProgress / stats.totalPharmacists) * 100 || 0)}%`],
    ["Not Started", stats.totalNotStarted, `${Math.round((stats.totalNotStarted / stats.totalPharmacists) * 100 || 0)}%`]
  ];

  autoTable(doc, {
    startY: yPos,
    head: headers,
    body: statsData,
    theme: 'striped',
    headStyles: { fillColor: [244, 164, 96] },
  });

  // @ts-ignore
  yPos = doc.lastAutoTable.finalY + 20;

  // Top Districts Table
  doc.text("Top Districts Performance", 14, yPos);
  yPos += 5;

  const districtHeaders = [["District", "Completed", "Pending", "Total"]];
  const districtData = stats.byDistrict.map(d => [
    d.name,
    d.Completed,
    d.Pending,
    d.Completed + d.Pending
  ]);

  autoTable(doc, {
    startY: yPos,
    head: districtHeaders,
    body: districtData,
    theme: 'grid',
    headStyles: { fillColor: [60, 60, 60] },
  });

  // Footer
  const pageCount = doc.getNumberOfPages();
  for (let i = 1; i <= pageCount; i++) {
    doc.setPage(i);
    doc.setFontSize(8);
    doc.text('United Pharmacy LMS Automation Tool', 14, doc.internal.pageSize.height - 10);
    doc.text(`Page ${i} of ${pageCount}`, pageWidth - 30, doc.internal.pageSize.height - 10);
  }

  doc.save("LMS_Executive_Summary.pdf");
};