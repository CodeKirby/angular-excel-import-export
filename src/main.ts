import 'zone.js/dist/zone';
import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { bootstrapApplication } from '@angular/platform-browser';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

@Component({
  selector: 'my-app',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
  <div style="text-align:center">
  <h1>Excel File Upload and Export</h1>
  <input type="file" (change)="onFileChange($event)" style="display:none" #fileInput>
  <button (click)="fileInput.click()">Upload Excel File</button>
  <button (click)="exportExcel()" [disabled]="!dataLoaded">Export Excel File</button>
  <button (click)="showFile()">Show file</button>
</div>

<div *ngIf="dataLoaded">
  <table>
    <tr *ngFor="let row of excelData; let rowIndex = index">
      <td *ngFor="let cell of row; let colIndex = index">
        <input [(ngModel)]="excelData[rowIndex][colIndex]" />
      </td>
    </tr>
  </table>
</div>
  `,
})
export class App {
  name = 'Angular';
  dataLoaded = false;
  excelData: any[][] = [];

  onFileChange(event: any) {
    const target: DataTransfer = <DataTransfer>event.target;

    if (target.files.length !== 1) {
      throw new Error('Cannot use multiple files');
    }

    const reader: FileReader = new FileReader();

    reader.onload = (e: any) => {
      const binaryString: string = e.target.result;
      const workbook: XLSX.WorkBook = XLSX.read(binaryString, {
        type: 'binary',
      });
      const sheetName: string = workbook.SheetNames[0];
      const worksheet: XLSX.WorkSheet = workbook.Sheets[sheetName];
      this.excelData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      this.dataLoaded = true;
    };

    reader.readAsBinaryString(target.files[0]);
  }

  exportExcel(): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(this.excelData);
    const workbook: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const wbout: ArrayBuffer = XLSX.write(workbook, {
      bookType: 'xlsx',
      type: 'array',
    });
    const file = new File([wbout], 'exported_data.xlsx', {
      type: 'application/octet-stream',
    });
    saveAs(file);
  }

  showFile() {
    const maxlength = this.excelData.reduce((maxArray, currentArray) => {
      return currentArray.length > maxArray.length ? currentArray : maxArray;
    }).length;

    this.excelData.forEach((element) => {
      while (element.length < maxlength) {
        element.push('');
      }
    });
  }
}

bootstrapApplication(App);
