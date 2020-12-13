using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;

namespace WebApplication2.Service
{
    /// <summary>
    /// エクセルCOMアプリケーションクラス
    /// </summary>
    public class ExcelApp : IDisposable {

        private Application App = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ExcelApp()
        {
            App = new Application();
        }

        /// <summary>
        /// 開いたエクセル内のBookを取得する
        /// </summary>
        /// <returns></returns>
        public Workbooks GetBooks()
        {
            return App.Workbooks;
        }

        /// <summary>
        /// アプリケーション解放処理
        /// </summary>
        public void Dispose()
        {
            if (App != null)
            {
                App.Quit();
                Marshal.ReleaseComObject(App);
            }
        }
    }

    /// <summary>
    /// エクセルCOMBooksクラス
    /// </summary>
    public class ExcelBooks : IDisposable
    {
        private Workbooks Books = null;
        private ExcelApp App = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ExcelBooks()
        {
            App = new ExcelApp();
            Books = App.GetBooks();
        }

        /// <summary>
        /// 対象のBookを開く
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public Workbook GetBook(string path)
        {
            return Books.Open(path);
        }

        /// <summary>
        /// 解放処理
        /// </summary>
        public void Dispose()
        {
            if (Books != null)
            {
                Books.Close();
                Marshal.ReleaseComObject(Books);
            }

            if (App != null)
            {
                App.Dispose();
            }
            
        }
    }

    /// <summary>
    /// エクセルBookクラス
    /// </summary>
    public class ExcelBook : IDisposable
    {
        private Workbook Book = null;
        private ExcelBooks Books = null;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path"> 対象のエクセルファイルのパス </param>
        public ExcelBook(string path)
        {
            Books = new ExcelBooks();
            Book = Books.GetBook(path);
        }

        /// <summary>
        /// シート一覧を取得する
        /// </summary>
        /// <returns></returns>
        public Sheets GetSheets()
        {
            return Book.Sheets;
        }

        /// <summary>
        /// 解放処理
        /// </summary>
        public void Dispose()
        {
            if (Book != null)
            {
                Book.Close();
                Marshal.ReleaseComObject(Book);
            }
            if (Books != null)
            {
                Books.Dispose();
            }
       
            
        }
    }

    /// <summary>
    /// エクセルSheetsクラス
    /// </summary>
    public class ExcelSheets : IDisposable
    {
        private Sheets Sheets = null;
        private ExcelBook Book = null;

        /// <summary>
        /// 解放処理
        /// </summary>
        /// <param name="path"></param>
        public ExcelSheets(string path)
        {
            Book = new ExcelBook(path);
            Sheets = Book.GetSheets();
        }

        /// <summary>
        /// シート名をすべて取得する
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetsName()
        {
            var sheetName = new List<string>();
            for (var i= 1;i <= Sheets.Count;i++)
            {
                using (var sheet = new ExcelSheet(Sheets[i]))
                {
                    sheetName.Add(sheet.GetSheetName());
                }
            }
            return sheetName;
        }

        /// <summary>
        /// シート内の値をすべて取得する
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public object[,] GetSheetValue(string sheetName)
        {
            using (var sheet = GetTargetSheet(sheetName))
            {
                return sheet.GetUseRamge();
            }
        }

        /// <summary>
        /// シート名から対象のシートを取得する
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ExcelSheet GetTargetSheet(string sheetName)
        {
            var idx = GetTargetSheetIdx(sheetName);
            if (idx == -1)
            {
                return null;
            }
            return (new ExcelSheet(Sheets[idx]));
        }

        /// <summary>
        /// 指定したシート名の該当する添え字を取得する
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public int GetTargetSheetIdx(string sheetName)
        {
            for (var i = 1;i <= Sheets.Count;i++)
            {
                using (var sheet = new ExcelSheet(Sheets[i]))
                {
                    if (sheet.GetSheetName() == sheetName)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        /// <summary>
        /// 解放処理
        /// </summary>
        public void Dispose()
        {
            Marshal.ReleaseComObject(Sheets);
            if (Book != null)
            {
                Book.Dispose();
            }
           
        }
    }

    /// <summary>
    /// エクセルSheetクラス
    /// </summary>
    public class ExcelSheet : IDisposable
    {
        private Worksheet Sheet;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="sheet"> 操作対象のシート </param>
        public ExcelSheet(Worksheet sheet)
        {
            Sheet = sheet;
        }

        /// <summary>
        /// シート内の値をobjectの配列型で取得する
        /// </summary>
        /// <returns> シートの値をすべて取得する </returns>
        public object[,] GetUseRamge()
        {
            using (var range = new ExcelRange(Sheet.UsedRange))
            {
                return range.GetUseRangeObj();
            }
        }

        /// <summary>
        /// 対応するシート名を取得する
        /// </summary>
        /// <returns> 対応するシートに該当するシート名を返す </returns>
        public string GetSheetName()
        {
            return Sheet.Name;
        }

        public void Dispose()
        {
            Marshal.FinalReleaseComObject(Sheet);
        }
    }

    /// <summary>
    /// エクセルRangeクラス
    /// </summary>
    public class ExcelRange : IDisposable
    {
        private Range Range;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="range"> 操作対象の範囲セル </param>
        public ExcelRange(Range range)
        {
            Range = range;
        }

        /// <summary>
        /// 指定した範囲の値を戻す
        /// </summary>
        /// <returns></returns>
        public object[,] GetUseRangeObj()
        {
            return Range.Value;
        }

        /// <summary>
        /// 終了処理
        /// </summary>
        public void Dispose()
        {
            Marshal.FinalReleaseComObject(Range);
        }
    }


    public class ExlceCom
    {
        // エクセルファイルの情報を取得して特定の値のみ戻す
        public void ReadExcelFile(string filePath,string sheetName)
        {
            object[,] value = PullExcelUserRange(filePath,sheetName);

        }

        /// <summary>
        /// 指定したBookの該当するシート名にある情報を取得
        /// 注意: あくまでWorkSheetのみ対応
        /// </summary>
        /// <param name="filePath"> エクセルのファイルパス </param>
        /// <param name="sheetName"> エクセルのシート名 </param>
        /// <returns></returns>
        public object[,] PullExcelUserRange(string filePath, string sheetName)
        {
            // シートを開く
            using (var sheets = new ExcelSheets(filePath))
            {
                // 開いたBook内の対象のシートに該当する値をすべて取得する
                return sheets.GetSheetValue(sheetName);
            }
        }
    }
}