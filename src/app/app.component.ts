import { Component, VERSION } from '@angular/core';
import * as XLSX from 'xlsx';

@Component({
  selector: 'my-app',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  fileName = 'ExcelSheet.xlsx';
  users: any = [
    { Id: 1, Name: 'Sahil', LastName: 'Manglani', Email: '123@gmail.com' },
    { Id: 2, Name: 'Savita', LastName: 'Manglani', Email: '123@gmail.com' },
    { Id: 3, Name: 'Bhavik', LastName: 'Manglani', Email: '123@gmail.com' },
    { Id: 4, Name: 'Ram', LastName: 'Manglani', Email: '123@gmail.com' },
    { Id: 5, Name: 'Shyam', LastName: 'Manglani', Email: '123@gmail.com' },
  ];
  exportexcel() {
    let element = document.getElementById('excel-table');
    const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(element);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }
}
