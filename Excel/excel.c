        /// <summary>
        /// Application�`Sheets�܂ł���� (�����Books�܂ł͌ʂɎg�p���Ȃ�����Sheets�ɂ܂Ƃ߂Ďd�l)
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
            /// �V�[�g�����w�肵�đΉ�����C���f�b�N�X���擾����
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
            /// �w�肵���C���f�b�N�X�̃V�[�g���擾����
            /// </summary>
            /// <param name="idx"></param>
            /// <returns></returns>
            public Worksheet GetSheet(int idx)
            {
                return Sheets[idx];
            }

            /// <summary>
            /// �������
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
        /// �V�[�g�P�̂̃��b�p�[�N���X
        /// </summary>
        public class ESheet : IDisposable
        {
            public Worksheet Sheet { get; set; } = null;

            public ESheet(Worksheet sheet)
            {
                Sheet = sheet;
            }

            /// <summary>
            /// �V�[�g�����擾����
            /// </summary>
            /// <returns></returns>
            public string GetSheetName()
            {
                return Sheet.Name;
            }

            /// <summary>
            /// �������
            /// </summary>
            public void Dispose()
            {
                Marshal.FinalReleaseComObject(Sheet);
            }
        }

        /// <summary>
        /// Range���b�p�[�N���X
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
        /// �G�N�Z���t�@�C���̎w�肵���V�[�g��2�����z��ɓǂݍ���.
        /// </summary>
        /// <param name="filePath">�G�N�Z���t�@�C���̃p�X</param>
        /// <param name="sheetIndex">�V�[�g�̔ԍ� (1, 2, 3, ...)</param>
        /// <param name="startRow">�ŏ��̍s (>= 1)</param>
        /// <param name="startColmn">�ŏ��̗� (>= 1)</param>
        /// <param name="lastRow">�Ō�̍s</param>
        /// <param name="lastColmn">�Ō�̗�</param>
        /// <returns>�V�[�g�����i�[����2���������z��. �������t�@�C���ǂݍ��݂Ɏ��s�����Ƃ��ɂ� null.</returns>
        public void Read(string filePath, int sheetIndex,
                              int startRow, int startColmn,
                              int lastRow, int lastColmn)
        {
            ESheets sheets = null;

            // �t�@�C���p�X
            string filaPath = "C:\\Users\\Owner\\Desktop\\workPro\\C#\\Resroce\\test.xlsx";

            // �t�@�C�����݃`�F�b�N
            if (!File.Exists(filaPath))
            {
                return;
            }

            // �V�[�g�܂ł�W�J 
            using (sheets = new ESheets(filaPath))
            {
                // �V�[�g��W�J����
                string sheetName = "test";
                int sheetIdx = sheets.GetSheetIdx(sheetName);

                // �V�[�g�̑��݃`�F�b�N
                if (sheetIdx == -1)
                {
                    return;
                }

                using (ESheet sheet = new ESheet(sheets.GetSheet(sheetIdx)))
                {
                    object[,] test = null;

                    // �g�p�ς݃Z�������ׂĎ擾����
                    using (ERange range = new ERange(sheet.Sheet))
                    {
                        test = range.GetRangeArray();
                    }
                }
            }

                //try
                //{
                   

                //    // Book�܂ł��J��
                //    mApp = new ESheets("");

                //    using (ESheets sheets = new ESheets(mWorkBook))
                //    {

                //    }

                //    // ���[�N�u�b�N���J��
                //    //Open(filePath, out mApp, out mWorkBooks, out mWorkBook);
                //}
                //catch (Exception ex)
                //{

                //}
                //finally
                //{
                //    // ����Ώہ@������:�K���Ō�ɊJ�������̂���������悤��
                //    mApp.Dispose();

                //    // ���[�N�u�b�N�ƃG�N�Z���̃v���Z�X�����
                //    //Close(mApp, mWrokBooks, mWorkBook);
                //}

            

            //Worksheet sheet = mWorkBook.Sheets[sheetIndex];
            //sheet.Select();

            //var arrOut = new ArrayList();

            //for (int r = startRow; r <= lastRow; r++)
            //{
            //    // ��s�ǂݍ���
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

            //// ���[�N�V�[�g�����
            //Marshal.ReleaseComObject(sheet);
            //sheet = null;

            //// ���[�N�u�b�N�ƃG�N�Z���̃v���Z�X�����
            //Close(mApp, mWrokBooks, mWorkBook);

            return ;
        }
