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

  // @ViewChild('tableDom2') tableDom!: ElementRef;





  constructor() { }



  excel(table: HTMLTableElement) {

    // console.log(table);
    // const table = tableDOM.nativeElement as HTMLTableElement;

    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');
    // let columns: Partial<Excel.Column>[] = [];


    let start_row: number = 1;
    let end_row: number = 1;

    const table_arr: Array<Array<number>> = [];

    // tr
    for (let i = 0; i < table.rows.length; i++) {

      start_row = i + 1;
      end_row = start_row;
      // console.log(`tr start_row=${start_row} , end_row=${end_row}`);

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
        end_row = start_row;

        const cell = row.cells[j];
        let cell_index: number = j;
        let text: string | number | Date = cell.innerText;





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

            // 開始欄
            if (table_arr[i][k + 1] === undefined) {

              table_arr[i][k + 1] = k + 1;
              cell_index = k + 1;
              start_cell = cell_index + 1;
              end_cell = start_cell;
              break;
            }

          }





        }









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


          // const target_call = j + cell.colSpan - 1;
          const target_call = start_cell - 1 + cell.colSpan - 1;
          // console.log(`target_call=${target_call}`);

          for (let k = start_cell; k <= target_call; k++) {
            table_arr[i][k] = -1;

            // console.log(`text=${text} , i=${i} , k=${k}`);


            if (cell.rowSpan > 1) {

              const target_row = i + cell.rowSpan - 1;
              // console.log(`cell.rowSpan=${cell.rowSpan} , target_row=${target_row}`);

              for (let l = i + 1; l <= target_row; l++) {


                if (table_arr[l][k] === undefined) {
                  // console.log(`l=${l} , k=${k}`);

                  // table_arr.push([]);
                  table_arr[l][k] = -1;
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
          // console.log(`cell.rowSpan=${cell.rowSpan} , start_row=${start_row} , end_row=${end_row}`);
          end_row = end_row + cell.rowSpan - 1;
        }

        // 合併儲存格
        if (cell.colSpan > 1 || cell.rowSpan > 1) {
          // console.log(`start_row=${start_row}, start_cell=${start_cell}, end_row=${end_row}, end_cell=${end_cell}`);
          worksheet.mergeCells(start_row, start_cell, end_row, end_cell);
        }



        // 設定標題欄
        // if (i === 0) {
        //   columns.push({ header: text, key: text });  // , width: 10 
        // }

        // 判斷型態
        if (Number(text.replace(/,/g, ''))) {

          text = Number(text.replace(/,/g, ''));

        }
        

        worksheet.getCell(start_row, start_cell).value = text;




        // style
        // backround color
        let fbColor: string = this.rgbaString2Hexargb(getComputedStyle(cell).backgroundColor);
        fbColor = fbColor.length === 0 ? 'ffffff' : fbColor;

        worksheet.getCell(start_row, start_cell).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: `${fbColor}` },
          bgColor: { argb: `ffffffff` }
        };


        // border 
        worksheet.getCell(start_row, start_cell).border = {
          top: { style: 'thin', color: { argb: this.rgbaString2Hexargb(getComputedStyle(cell).borderTopColor) } },
          left: { style: 'thin', color: { argb: this.rgbaString2Hexargb(getComputedStyle(cell).borderLeftColor) } },
          bottom: { style: 'thin', color: { argb: this.rgbaString2Hexargb(getComputedStyle(cell).borderBottomColor) } },
          right: { style: 'thin', color: { argb: this.rgbaString2Hexargb(getComputedStyle(cell).borderRightColor) } }
        };


        // font
        worksheet.getCell(start_row, start_cell).font = {
          // name: ,
          color: { argb: this.rgbaString2Hexargb(getComputedStyle(cell).color) },
          // family: getComputedStyle(cell).fontFamily,
          size: Number(getComputedStyle(cell).fontSize.replace(/[^\d+]./g, '')),
          // italic: true
        };





        // 設定下一個開始欄位
        end_cell++;
        start_cell = end_cell;
        // start_row++;
        // rowValues[j] = text;
      }

      // console.log(table_arr);
    }


    // console.log(table_arr);

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


    // worksheet.getColumn(2).numFmt = '#,##0.00;[Red]\-#,##0.00';


    workbook.xlsx.writeBuffer().then((buffer) => {
      const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const fileExtension = '.xlsx';
      const blob = new Blob([buffer], { type: fileType });
      saveAs(blob, 'exceljs' + fileExtension);
    });


  }




  rgbaString2Hexargb(str: string): string {
    // str=
    // rgb(200,210,238)
    // rgba(0,0,0,0)

    str = str.replace(/ /g, '');
    // const regex = /(\d+),(\d+),(\d+)/;
    const regex = /(\d+),(\d+),(\d+),?(\d)?/;
    let m;

    if ((m = regex.exec(str)) !== null) {

      if (m[4] !== undefined && Number(m[4]) === 0) {
        return 'ffffffff';
      }

      return '00' + this.rgb2hex(Number(m[1]), Number(m[2]), Number(m[3]));
    }

    return '';
  }



  rgb2hex(r: number, g: number, b: number): string {
    var rgb = (r << 16) | (g << 8) | b
    // return '#' + rgb.toString(16) // #80c0
    // return '#' + (0x1000000 + rgb).toString(16).slice(1) // #0080c0
    // or use [padStart](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String/padStart)
    return rgb.toString(16).padStart(6, '0');
  }



}
