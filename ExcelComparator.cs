using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace ExcelDiff
{
    class ExcelComparator
    {
        private string fileA;
        private string fileB;
        private int totalLines;

        public ExcelComparator(string a, string b)
        {
            this.fileA = a;
            this.fileB = b;
        }

        public void Compare(string sheetName, int column, ProgressBar progress)
        {
            this.totalLines = 2 * (Lines(sheetName, fileA) + Lines(sheetName, fileB));
            progress.Maximum = totalLines;
            progress.Value = 0;
            OnlyInA(sheetName, column, progress);
            OnlyInB(sheetName, column, progress);

        }
        private bool existsIn(List<object> row, List<List<object>> contentA)
        {
            foreach (List<object> candidate in contentA)
            {
                bool different = false;
                for (int i = 0; i < row.Count; ++i)
                {
                    if (row[i] == null)
                    {
                        if (candidate.Count < i)
                        {
                            continue;
                        }
                        if (candidate[i] != null)
                        {
                            different = true;
                            break;
                        }
                    }
                    else if (candidate.Count <= i)
                    {
                        throw new Exception("Nombre de colonnes différents ! (" + row.Count + " vs " + candidate.Count + ")");
                    }
                    else
                    if (!(row[i].Equals(candidate[i])))
                    {
                        different = true;
                        break;
                    }
                }
                if (!different)
                {
                    return true;
                }
            }
            return false;
        }

        public void OnlyInA(string sheetName, int column, ProgressBar progress)
        {
            OnlyIn(sheetName, column, fileB, fileA, progress);
        }

        public void OnlyInB(string sheetName, int column, ProgressBar progress)
        {
            OnlyIn(sheetName, column, fileA, fileB, progress);
        }

        public void OnlyIn(string sheetName, int column, string fileA, string fileB, ProgressBar progress)
        {
            HashSet<object> contentA = Keys(sheetName, column, fileA, progress);

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            int rCnt;
            int rw = 0;
            int cl = 0;


            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileB, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(sheetName))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + fileB);
                return;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            if (cl < column + 1)
            {
                MessageBox.Show("Pas de colonne " + (column + 1) + " dans " + fileB);

                return;
            }


            List<int> removed = new List<int>();
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {

                var cell = (range.Cells[rCnt, column + 1] as Excel.Range).Value2;

                if (contentA.Contains(cell))
                {
                    // same line remove it !
                    //for (cCnt = 1; cCnt <= cl; cCnt++)
                    //{
                    //    (range.Cells[rCnt, cCnt] as Excel.Range).Value2 = "";
                    //    //range[rCnt, cCnt].Delete();
                    //}
                    removed.Add(rCnt);
                }

                progress.Value += 1;
            }

            foreach (var line in removed.Reverse<int>())
            {
                range.Rows[line].Delete();
            }

            xlWorkBook.SaveCopyAs(Path.GetDirectoryName(fileB) + @"\Dans-" + Path.GetFileName(fileB) + "-mais-pas-dans-" + Path.GetFileName(fileA));

            xlWorkBook.Close(false, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
        /*
        private List<List<object>> Content(string file, ProgressBar progress)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            List<List<object>> content = new List<List<object>>();

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(this.sheetName))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + file);

                return content;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            progress.Maximum += rw;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                List<object> row = new List<object>();
                content.Add(row);
                for (cCnt = 1; cCnt <= cl; cCnt++)
                {
                    row.Add((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                }
                progress.Value += 1;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return content;
        }
        */
        private HashSet<object> Keys(string sheetName, int column, string file, ProgressBar progress)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            string str;
            int rCnt;
            int cCnt;
            int rw = 0;
            int cl = 0;

            HashSet<object> content = new HashSet<object>();

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
            {
                if (sheet.Name.Equals(sheetName))
                {
                    xlWorkSheet = sheet;
                }
            }

            if (null == xlWorkSheet)
            {
                MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + file);

                return content;
            }

            range = xlWorkSheet.UsedRange;
            rw = range.Rows.Count;
            cl = range.Columns.Count;

            if (cl < column + 1)
            {
                MessageBox.Show("Pas de colonne " + (column + 1) + " dans " + file);

                return content;
            }

            //progress.Maximum += rw;


            for (rCnt = 1; rCnt <= rw; rCnt++)
            {
                content.Add((range.Cells[rCnt, column + 1] as Excel.Range).Value2);
                progress.Value += 1;
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
            return content;
        }

        private int Lines(string sheetName, string file)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    if (sheet.Name.Equals(sheetName))
                    {
                        xlWorkSheet = sheet;
                    }
                }

                if (null == xlWorkSheet)
                {
                    MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + file);

                    return -1;
                }
                range = xlWorkSheet.UsedRange;
                return range.Rows.Count;
            }
            finally
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public List<Column> Columns(string sheetName)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet = null;
            Excel.Range range;

            List<Column> ret = new List<Column>();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    if (sheet.Name.Equals(sheetName))
                    {
                        xlWorkSheet = sheet;
                    }
                }

                if (null == xlWorkSheet)
                {
                    MessageBox.Show("Pas d'onglet \"" + sheetName + "\" dans " + fileA);
                    return ret;
                }

                range = xlWorkSheet.UsedRange;
                int cols = range.Columns.Count;
                for (int i = 0; i < cols; ++i)
                {
                    ret.Add(new Column(i));
                }
                return ret;
            }
            finally
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                if (null!=xlWorkSheet) Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }

        public List<string> Sheets()
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;

            List<string> ret = new List<string>();
            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(fileA, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            try
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Worksheets)
                {
                    ret.Add(sheet.Name);
                }
                return ret;
            }
            finally
            {

                xlWorkBook.Close(true, null, null);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}
