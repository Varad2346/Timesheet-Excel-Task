import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import * as XLSX from 'xlsx';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css',
})
export class AppComponent {
  data: any[] = [];
  // Reading a file as string
  onFileUpload(event: any) {
    const target: DataTransfer = <DataTransfer>event.target;
    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }

    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      this.data = XLSX.utils.sheet_to_json(ws, { raw: false });
      console.log("Original data loaded:", this.data);
    };
    reader.readAsArrayBuffer(target.files[0]);
  }

  // Export function
  transformAndExport(): void {
   
    const dateMap = new Map<string, any>();
    const customerNameSet = new Set<string>();

    // Data Merging Using Date Field
    this.data.forEach((row: any) => {
      customerNameSet.add(row['Customer Name']);

      const efforts = (parseFloat(row.Hours || 0)) + (parseFloat(row.Minutes || 0) / 60);
      const dateKey = row['Date'];
      const existingEntry = dateMap.get(dateKey);

      if (existingEntry) {
        existingEntry['Efforts(Hours)'] += efforts;
        existingEntry['TASK DESCRIPTION'] += `, ${row['Task Description']}`;
      } else {
        dateMap.set(dateKey, {
          'DATE': dateKey,
          'DAY': new Date(dateKey).toLocaleDateString('en-us', { weekday: 'long' }),
          'Efforts(Hours)': efforts,
          'TASK DESCRIPTION': `${row['Task Description']}`,
        });
      }
    });

    const sortedEntries = Array.from(dateMap.values())
      .sort((a, b) => new Date(a.DATE).getTime() - new Date(b.DATE).getTime());

    let srNo = 1;
    let holidays = 0;
  
    const aggregatedData = sortedEntries.map((entry) => {
      const date = new Date(entry.DATE);
      const day = date.getDay(); 
      if (day === 6 || day === 7) {
        holidays++;
      }

      return {
        "Sr.No.": srNo++,
        ...entry
      };
    });

    // Description Split Code
    const finalSheetData: any[] = [];
    aggregatedData.forEach((item) => {
      const tasks = item['TASK DESCRIPTION']
        .split(',')
        .map((t: string) => t.trim()) 
        .filter((t: string) => t);  

      if (tasks.length === 0) {
        finalSheetData.push({ ...item, 'TASK DESCRIPTION': '' });
      } else {
        const firstTask = tasks.shift();
        finalSheetData.push({
          ...item,
          'TASK DESCRIPTION': `1. ${firstTask}`,
        });

        tasks.forEach((task: string, index: number) => {
          finalSheetData.push({
            'Sr.No.': '',
            'DATE': '',
            'DAY': '',
            'Efforts(Hours)': '',
            'TASK DESCRIPTION': `${index + 2}. ${task}`,
          });
        });
      }
    });

    // Customer Names List
    const customerNamesList = Array.from(customerNameSet).map((name, i) => `${i + 1}. ${name}`);
    const totalDays = aggregatedData.length;
    const workedDays = totalDays - holidays;

    // Upper Header Declaration
    const aoa: any[][] = [
      ["CUSTOMER NAME", "", customerNamesList[0] || '', "", "PROJECT MANAGER", "VARAD", "CALENDAR DAYS", totalDays],
      ["RESOURCE NAME", "", "AJIT", "", "APPROVER NAME", "VEDANT", "WEEKLY OFF/HOLIDAYS", holidays],
      ["FOR MONTH", "", "", "", "SUBMITTED BY", "AJIT", "WORKED DAYS", workedDays],
      ["ROLE", "", "FRONTEND", "", "SUBMISSION DATE", "", "LEAVES TAKEN", ""],
      ["", "", "", "", "", "", "OVERTIME DAYS", ""]
    ];

    // Customer List Formatting
    if (customerNamesList.length > 1) {
      const remainingNamesAOA = customerNamesList.slice(1).map(name => ["", "", name, "", "", "", ""]);
      aoa.splice(1, 0, ...remainingNamesAOA); 
    }

    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(aoa);

    // Merge Code Section
    ws['!merges'] = [];
    const headerRowCount = customerNamesList.length + 4;

    for (let i = 0; i < headerRowCount - 1; i++) { 
      ws['!merges'].push({ s: { r: i, c: 2 }, e: { r: i, c: 3 } });
    }

    for (let i = 0; i < headerRowCount-1; i++) {
      ws['!merges'].push({ s: { r: i, c: 0 }, e: { r: i, c: 1 } });
    }

    const dataHeaderRowIndex = 8; 
    const numRowsToMerge = 1 + finalSheetData.length; 
    for (let i = 0; i < numRowsToMerge; i++) {
      const rowIndex = dataHeaderRowIndex + i;
      ws['!merges'].push({
        s: { r: rowIndex, c: 4 },
        e: { r: rowIndex, c: 7 } 
      });
    }
    // Column Width
    ws['!cols'] = [
      { wch: 5 },  // A: Sr.No.
      { wch: 12 }, // B: DATE
      { wch: 12 }, // C: DAY
      { wch: 12 }, // D: Efforts
      { wch: 17 }, // E: TASK DESCRIPTION 
      { wch: 12 }, // F
      { wch: 20 }, // G
    ];

    // Adding JSON
    XLSX.utils.sheet_add_json(
      ws,
      finalSheetData, {
      origin: `A+${customerNamesList.length+6}`, 
      skipHeader: false
    }
    );

    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'CombinedReport.xlsx');
  }
}