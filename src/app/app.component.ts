import { Component, ElementRef, ViewChild } from '@angular/core';
import * as Excel from 'exceljs';
import { saveAs } from 'file-saver';
import { TableModel } from './interfaces/table-model';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'exceljs-demo';

  list: TableModel[] = [
    { string: 'aaa', number: 1000, date: '2021-01-01 11:00' },
    { string: 'bbb', number: 2000, date: '2021-01-02 22:00' },
    { string: 'ccc', number: -3000, date: '2021-01-03 13:00' },
  ];

  @ViewChild('tableDom') tableDom!: ElementRef;


  constructor() { }



  excel() {

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    // let columns: Partial<Excel.Column>[] = [];


    console.log(this.tableDom);

    const table = this.tableDom.nativeElement as HTMLTableElement;
    let start_row: number = 1;
    let end_row: number = 1;
    let rowspan: number = 1;
    const table_arr: Array<Array<number>> = [];

    // tr
    for (let i = 0; i < table.rows.length; i++) {

      start_row = i + 1;
      end_row = start_row;

      if (table_arr[i] === undefined) {
        table_arr.push([]);
      }

      // console.log(table_arr);

      // tr
      const row = table.rows[i] as HTMLTableRowElement;
      // console.log(row);
      // const rowValues = [];
      let start_cell: number = 1;
      let end_cell: number = 1;

      // th, td
      for (let j = 0; j < row.cells.length; j++) {
        // console.log(`for start_cell=${start_cell}`);

        const cell = row.cells[j];
        let cell_index: number = j;
        let text: string | number | Date = cell.innerText;
        // console.log(cell, getComputedStyle(cell).backgroundColor);
        // console.log(cell, cell.style);
        // console.log(cell.style.backgroundColor.toString());


        // if (table_arr[i][j] === undefined) {
        //   table_arr[i].push(1);
        // } else {

        // }


        // 合併儲存格


        if (table_arr[i][j] === undefined) {
          // console.log(`i=${i} , j=${j} , table_arr[i][j]=${table_arr[i][j]}`);
          // table_arr.push([]);
          table_arr[i][j] = j;

          // if (i === 7) {
          //   console.log(`i=${i}  , j=${j}`);
          // }

        } else {

          // console.log(`i=${i} , j=${j} , table_arr[i][j]=${table_arr[i][j]}`);


          // 此欄已經有值，就往後找
          for (let k = j; k > -1; k++) {

            // console.log(`table_arr[i][j + 1] = ${table_arr[i][j + 1]}`);

            // 開始欄
            if (table_arr[i][j + 1] === undefined) {
              // console.log(`[j + 1]`);
              table_arr[i][j + 1] = j + 1;
              cell_index = j + 1;
              start_cell = cell_index + 1;
              end_cell = start_cell;
              break;
            }

          }





          // const now_length = table_arr[i].length;

          // for (let k = j; k < now_length; k++) {

          //   // console.log(`i=${i} , k=${k}`);

          //   if (table_arr[i][k] === -1) {

          //     start_cell++;
          //     end_cell = start_cell;
          //     // console.log(`now_length=${now_length} , table_arr[i=${i}][k=${k}] ,start_cell=${start_cell}`);
          //   }

          // }





        }







        // //                               0  
        // const target_cell = table_arr[i].length - 1 + cell.colSpan;
        // const now_length = table_arr[i].length;

        // // console.log(table_arr);
        // for (let k = j; k <= target_cell; k++) {

        //   // console.log(`i=${i} , k=${k} , table_arr[i][k]=${table_arr[i][k]}`);
        //   // console.log(`i=${i} , k=${k} , table_arr[i][k]=${table_arr[i][k]}`);

        //   if (table_arr[i][k] === undefined) {
        //     // console.log(`i=${i} , k=${k} , table_arr[i][k]=${table_arr[i][k]}`);
        //     // table_arr.push([]);
        //     table_arr[i][k] = k;

        //     if (i === 7) {
        //       console.log(`i=${i}  , k=${k}`);
        //     }

        //   }


        // }


        // for (let k = j; k < now_length; k++) {

        //   // console.log(`i=${i} , k=${k}`);

        //   if (table_arr[i][k] === -1) {

        //     start_cell++;
        //     end_cell = start_cell;
        //     console.log(`now_length=${now_length} , table_arr[i=${i}][k=${k}] ,start_cell=${start_cell}`);
        //   }

        // }

        // console.log(`i=${i} , j=${j} , start_cell=${start_cell}`);







        if (cell.rowSpan > 1) {

          const target_row = i + cell.rowSpan - 1;

          for (let k = i + 1; k <= target_row; k++) {
            // console.log(`k=${k} , start_cell=${start_cell}`);

            if (table_arr[k] === undefined) {
              table_arr[k] = [];
              // table_arr[k].push(1);
            }

            if (table_arr[k][start_cell - 1] === undefined) {
              // table_arr.push([]);
              table_arr[k][start_cell - 1] = -1;
              // console.log(`[${k}][${start_cell - 1}] , k=${k} , start_cell=${start_cell - 1}`);
            }

          }

        }



        if (cell.colSpan > 1) {


          const target_call = j + cell.colSpan - 1;
          // console.log(`target_call=${target_call}`);

          for (let k = j + 1; k <= target_call; k++) {
            table_arr[i][k] = -1;


            if (cell.rowSpan > 1) {

              const target_row = i + cell.rowSpan - 1;
              console.log(`cell.rowSpan=${cell.rowSpan} , target_row=${target_row}`);

              for (let l = i + 1; l <= target_row; l++) {

                
                if (table_arr[l][k] === undefined) {
                  console.log(`l=${l} , k=${k}`);
                  
                  // table_arr.push([]);
                  // table_arr[l][k] = -1;
                  // console.log(`[${k}][${start_cell - 1}] , k=${k} , start_cell=${start_cell - 1}`);
                }



              }



            }


            // console.log(`forforforforforfor  i=${i} , k=${k}`);
          }


        }








        if (cell.colSpan > 1) {
          end_cell += cell.colSpan - 1;
        }

        if (cell.rowSpan > 1) {
          end_row += cell.rowSpan - 1;
        }

        if (cell.colSpan > 1 || cell.rowSpan > 1) {
          // console.log(`start_row=${start_row}, start_cell=${start_cell}, end_row=${end_row}, end_cell=${end_cell}`);
          worksheet.mergeCells(start_row, start_cell, end_row, end_cell);
        }



        // if (cell.rowSpan > 1) {

        // const target_row = j + cell.rowSpan - 1;

        // for (let k = j; k < target_row; k++) {
        //   if (table_arr[i] === undefined) {
        //     table_arr.push([]);
        //     table_arr[k].push(1);
        //   }
        // }







        // 設定標題欄
        // if (i === 0) {
        //   columns.push({ header: text, key: text });  // , width: 10 
        // }

        // 判斷型態
        if (Number(text.replace(/,/g, ''))) {

          text = Number(text.replace(/,/g, ''));

        }


        // console.log(text, cell.colSpan);

        // console.log(rowID, cellID);


        worksheet.getCell(start_row, start_cell).value = text;

        // 設定下一個開始欄位
        end_cell++;
        start_cell = end_cell;
        // start_row++;
        // rowValues[j] = text;
      }

      // if (i === 0) {
      //   worksheet.columns = columns;
      // }

      // worksheet.addRow(rowValues);


      // console.log(table_arr);
    }


    console.log(table_arr);

    // A
    // worksheet.getColumn(1).fill = {
    //   type: 'pattern',
    //   pattern: 'solid',
    //   fgColor: { argb: 'FFFFFF00' },
    //   bgColor: { argb: '00FFFF00' }
    // };

    // worksheet.getCell(1, 2).fill = {
    //   type: 'pattern',
    //   pattern: 'solid',
    //   fgColor: { argb: 'FFFFFF00' },
    //   bgColor: { argb: '00FFFF00' }
    // };


    worksheet.getColumn(2).numFmt = '#,##0.00;[Red]\-#,##0.00';


    // workbook.xlsx.writeBuffer().then((buffer) => {
    //   const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    //   const fileExtension = '.xlsx';
    //   const blob = new Blob([buffer], { type: fileType });
    //   saveAs(blob, 'exceljs' + fileExtension);
    // });


  }

}
