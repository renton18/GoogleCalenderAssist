using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace GoogleCalenderAssist
{
    public class EXCEL
    {
        public string fname;
        public Application ExcelApp;
        public Workbook wb;

        public EXCEL()
        {
            ExcelApp = new Application();
            ExcelApp.Visible = false;
        }

        /// <summary>
        /// 新規作成
        /// </summary>
        public void Make()
        {            
            wb = ExcelApp.Workbooks.Add();
        }

        /// <summary>
        /// 開く
        /// </summary>
        public void Open(string fname)
        {
            ExcelApp.Visible = false;

            this.fname = fname;

            //エクセルファイルをオープンする
            wb = (Workbook)(ExcelApp.Workbooks.Open(
              System.Environment.CurrentDirectory + "/" + fname,  // オープンするExcelファイル名
              Type.Missing, // （省略可能）UpdateLinks (0 / 1 / 2 / 3)
              Type.Missing, // （省略可能）ReadOnly (True / False )
              Type.Missing, // （省略可能）Format
                // 1:タブ / 2:カンマ (,) / 3:スペース / 4:セミコロン (;)
                // 5:なし / 6:引数 Delimiterで指定された文字
              Type.Missing, // （省略可能）Password
              Type.Missing, // （省略可能）WriteResPassword
              Type.Missing, // （省略可能）IgnoreReadOnlyRecommended
              Type.Missing, // （省略可能）Origin
              Type.Missing, // （省略可能）Delimiter
              Type.Missing, // （省略可能）Editable
              Type.Missing, // （省略可能）Notify
              Type.Missing, // （省略可能）Converter
              Type.Missing, // （省略可能）AddToMru
              Type.Missing, // （省略可能）Local
              Type.Missing  // （省略可能）CorruptLoad
            ));
        }

        /// <summary>
        /// 閉じる
        /// </summary>
        public void Close()
        {
            // Book解放
            wb.Close(false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            // Excelアプリケーションを解放
            ExcelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
        }

        /// <summary>
        /// 名前を付けて保存
        /// </summary>
        public void Save(string fname)
        {
            wb.SaveAs(System.Environment.CurrentDirectory + "/" + fname);
        }

        /// <summary>
        /// 新しいシート作成
        /// </summary>
        public void MakeNewSheet(string sheetname)
        {
            Worksheet sheet;
            //最後尾に追加
            sheet = (Worksheet)ExcelApp.Worksheets.Add(After : wb.Worksheets[wb.Worksheets.Count]);  
            sheet.Name = sheetname;

            //Sheet解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
        }

        /// <summary>
        /// シート編集
        /// </summary>
        public void EditSheet(int year)
        {
            //シート作成
            for (int i = 1; i < 13; i++)
            {
                if (i < 4)
                {
                    wb.Worksheets[i].Name = i + "月";
                }
                else
                {
                    MakeNewSheet(i + "月");
                }
                //表示倍率
                ExcelApp.ActiveWindow.Zoom = 70;
            }

            //月設定
            for (int i = 1; i < 13; i++)
            {
                //月シートアクティブ
                Worksheet sheet = (Worksheet)wb.Sheets[i];

                //日数取得
                int daycount = DateTime.DaysInMonth(year, i);

                //日設定
                int col;
                int row;
                col = 1;
                row = 2;

                for(int j = 1; j <= daycount; j++)
                {
                    Range ran;
                    //ヘッダー
                    sheet.Cells[1, 1] = "日付";
                    sheet.Cells[1, 2] = "時間";
                    sheet.Cells[1, 3] = "マイカレンダー1　大阪オフィス";
                    sheet.Cells[1, 4] = "場所";
                    sheet.Cells[1, 5] = "時間";
                    sheet.Cells[1, 6] = "マイカレンダー2　東京オフィス";
                    sheet.Cells[1, 7] = "場所";

                    //列幅
                    sheet.Cells[1, 1].ColumnWidth = 10;
                    sheet.Cells[1, 2].ColumnWidth = 10;
                    sheet.Cells[1, 3].ColumnWidth = 30;
                    sheet.Cells[1, 4].ColumnWidth = 10;
                    sheet.Cells[1, 5].ColumnWidth = 10;
                    sheet.Cells[1, 6].ColumnWidth = 30;
                    sheet.Cells[1, 7].ColumnWidth = 10;

                    //ヘッダー格子
                    ran = ExcelApp.Range[sheet.Cells[1, 1], sheet.Cells[1, 7]];
                    ran.Borders.LineStyle = XlLineStyle.xlContinuous;
                    //中央揃え
                    ran.HorizontalAlignment = XlHAlign.xlHAlignCenter;     

                    //セル結合
                    ran = ExcelApp.Range[sheet.Cells[row, col],sheet.Cells[row+23, col]];
                    ran.Merge();

                    //罫線
                    for (int k = 1; k < 8; k++)
                    {
                        ran = ExcelApp.Range[sheet.Cells[row, k], sheet.Cells[row + 23, k]];
                        ran.BorderAround();
                    }
                    sheet.Cells[row, col] = i + "月" + j + "日";

                    //中央揃え
                    sheet.Cells[row, col].HorizontalAlignment = XlHAlign.xlHAlignCenter;  

                    row = row + 24;
                }

                //Sheet解放
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);                
            }
        }

        /// <summary>
        /// データ入力
        /// </summary>
        public void Write(int month, int day, string time, string place, string content, int calID)
        {
            //月シートアクティブ
            Worksheet sheet = (Worksheet)wb.Sheets[month];

            //時間よりセルの特定
            string[] timerow = time.Split(':');
            //行特定
            int row;
            row = 1 + (day -1) * 24 + int.Parse(timerow[0]);

            //列特定
            if (calID == 1)
            {
                sheet.Cells[row, 2] = time;
                sheet.Cells[row, 3] = content;
                sheet.Cells[row, 4] = place;
            }else
            {
                sheet.Cells[row, 5] = time;
                sheet.Cells[row, 6] = content;
                sheet.Cells[row, 7] = place;
            }
        }
    }
}
