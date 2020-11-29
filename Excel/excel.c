        /// <summary>
        /// Application～Sheetsまでを解放 (今回はBooksまでは個別に使用しないためSheetsにまとめて仕様)
        /// </summary>
        public class ESheets : IDisposable
        {
            public Application App { get; set; } = null;
            public Workbooks Books { get; set; } = null;
            public Workbook Book { get; set; } = null;
            public Sheets Sheets { get; set; } = null;

            public ESheets(string path)
            {
                App = new Application();
                App.Visible = false;

                Books = App.Workbooks;
                Book = Books.Open(path);

                Sheets = Book.Sheets;
            }
            
            /// <summary>
            /// シート名を指定して対応するインデックスを取得する
            /// </summary>
            /// <param name="sheetName"></param>
            /// <returns></returns>
            public int GetSheetIdx(string sheetName)
            {
                for (int i = 1; i <= Sheets.Count;i++)
                {
                    using (ESheet sheet = new ESheet(Sheets[i]))
                    {
                        if (sheet.GetSheetName() == sheetName) { return i; }
                    }
                }

                return -1;
            }

            /// <summary>
            /// 指定したインデックスのシートを取得する
            /// </summary>
            /// <param name="idx"></param>
            /// <returns></returns>
            public Worksheet GetSheet(int idx)
            {
                return Sheets[idx];
            }

            /// <summary>
            /// 解放処理
            /// </summary>
            public void Dispose()
            {
                Marshal.FinalReleaseComObject(Sheets);

                if (Book != null)
                {
                    Book.Close();
                }
                Marshal.FinalReleaseComObject(Book);


                if (Books != null)
                {
                    Books.Close();
                }
                Marshal.FinalReleaseComObject(Books);

                if (App != null)
                {
                    App.Quit();
                }
                Marshal.FinalReleaseComObject(App);
            }
        }

        /// <summary>
        /// シート単体のラッパークラス
        /// </summary>
        public class ESheet : IDisposable
        {
            public Worksheet Sheet { get; set; } = null;

            public ESheet(Worksheet sheet)
            {
                Sheet = sheet;
            }

            /// <summary>
            /// シート名を取得する
            /// </summary>
            /// <returns></returns>
            public string GetSheetName()
            {
                return Sheet.Name;
            }

            /// <summary>
            /// 解放処理
            /// </summary>
            public void Dispose()
            {
                Marshal.FinalReleaseComObject(Sheet);
            }
        }

        /// <summary>
        /// Rangeラッパークラス
        /// </summary>
        public class ERange : IDisposable
        {
            Range range { get; set; } = null;

            public ERange(Worksheet sheet)
            {
                range = sheet.UsedRange;
            }

            public Object[,] GetRangeArray()
            {
                return range.Value;
            }

            public void Dispose()
            {
                Marshal.FinalReleaseComObject(range);
            }
        }



        public Excel()
        {

        }

        /// <summary>
        /// エクセルファイルの指定したシートを2次元配列に読み込む.
        /// </summary>
        /// <param name="filePath">エクセルファイルのパス</param>
        /// <param name="sheetIndex">シートの番号 (1, 2, 3, ...)</param>
        /// <param name="startRow">最初の行 (>= 1)</param>
        /// <param name="startColmn">最初の列 (>= 1)</param>
        /// <param name="lastRow">最後の行</param>
        /// <param name="lastColmn">最後の列</param>
        /// <returns>シート情報を格納した2次元文字配列. ただしファイル読み込みに失敗したときには null.</returns>
        public void Read(string filePath, int sheetIndex,
                              int startRow, int startColmn,
                              int lastRow, int lastColmn)
        {
            ESheets sheets = null;

            // ファイルパス
            string filaPath = "C:\\Users\\Owner\\Desktop\\workPro\\C#\\Resroce\\test.xlsx";

            // ファイル存在チェック
            if (!File.Exists(filaPath))
            {
                return;
            }

            // シートまでを展開 
            using (sheets = new ESheets(filaPath))
            {
                // シートを展開する
                string sheetName = "test";
                int sheetIdx = sheets.GetSheetIdx(sheetName);

                // シートの存在チェック
                if (sheetIdx == -1)
                {
                    return;
                }

                using (ESheet sheet = new ESheet(sheets.GetSheet(sheetIdx)))
                {
                    object[,] test = null;

                    // 使用済みセルをすべて取得する
                    using (ERange range = new ERange(sheet.Sheet))
                    {
                        test = range.GetRangeArray();
                    }
                }
            }

                //try
                //{
                   

                //    // Bookまでを開く
                //    mApp = new ESheets("");

                //    using (ESheets sheets = new ESheets(mWorkBook))
                //    {

                //    }

                //    // ワークブックを開く
                //    //Open(filePath, out mApp, out mWorkBooks, out mWorkBook);
                //}
                //catch (Exception ex)
                //{

                //}
                //finally
                //{
                //    // 解放対象　※注意:必ず最後に開いたものから解放するように
                //    mApp.Dispose();

                //    // ワークブックとエクセルのプロセスを閉じる
                //    //Close(mApp, mWrokBooks, mWorkBook);
                //}

            

            //Worksheet sheet = mWorkBook.Sheets[sheetIndex];
            //sheet.Select();

            //var arrOut = new ArrayList();

            //for (int r = startRow; r <= lastRow; r++)
            //{
            //    // 一行読み込む
            //    var row = new ArrayList();
            //    for (int c = startColmn; c <= lastColmn; c++)
            //    {
            //        var cell = sheet.Cells;
            //        var cellValue = cell[r, c];

            //        if (cellValue == null || cellValue.Value == null) { row.Add(""); }

            //        row.Add(cellValue.Value);

            //        Marshal.ReleaseComObject(cell);
            //        Marshal.ReleaseComObject(cellValue);
            //    }

            //    arrOut.Add(row);
            //}

            //// ワークシートを閉じる
            //Marshal.ReleaseComObject(sheet);
            //sheet = null;

            //// ワークブックとエクセルのプロセスを閉じる
            //Close(mApp, mWrokBooks, mWorkBook);

            return ;
        }
