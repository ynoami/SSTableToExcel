using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SSTableToExcel.Models
{
    public class SSTableExcelBook
    {
        private List<string> _tables = new List<string>();
        private XLWorkbook _book = new XLWorkbook();

        /// <summary>
        /// テーブル情報Excelシートを追加
        /// </summary>
        /// <param name="records">テーブル情報</param>
        /// <param name="tableName">テーブル名</param>
        /// <param name="columnNames">テーブル列名の並び</param>
        public void AppendExcelSheet(List<dynamic> records, string tableName, string[] columnNames)
        {
            Progress.Set(tableName);

            // テーブルをインデックスに追記
            AppendIndex(tableName, records.Count);

            // もし、レコード数が0ならシートを追加しない
            if (records.Count == 0) return;

            // テーブルシートを作成
            var sheet = _book.AddWorksheet(tableName);

            // シートにヘッダを書き込み (1行目)
            for (int index = 0; index < columnNames.Length; index++)
            {
                var headCell = sheet.Cell(1, index + 1);
                headCell.Value = columnNames[index];
                headCell.Style.Alignment.TextRotation = 45;
            }

            // シートに全レコード書き込み (2行目以降)
            for (int rowIndex = 0; rowIndex < records.Count; rowIndex++)
            {
                if ((rowIndex % 50) == 0)
                {
                    Progress.AppendProgress();
                }

                var columnValues = (records[rowIndex] as IDictionary<string, object>).Select(_ => _.Value).ToArray();
                for (int columnIndex = 0; columnIndex < columnNames.Length; columnIndex++)
                {
                    var fieldCell = sheet.Cell(rowIndex + 2, columnIndex + 1);
                    fieldCell.Value = columnValues[columnIndex];
                    fieldCell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                }
            }

            Progress.Clear();
        }

        /// <summary>
        /// Excelファイルを保存する
        /// </summary>
        /// <param name="fileName">保存先ファイル名</param>
        public void SaveAs(string fileName)
        {
            Progress.Set("Excelファイルを保存中です");
            _book.SaveAs(fileName);
            Progress.Clear();
        }

        /// <summary>
        /// テーブル情報をインデックスに追加する
        /// </summary>
        /// <param name="tableName">テーブル名</param>
        /// <param name="count"></param>
        private void AppendIndex(string tableName, int count)
        {
            IXLWorksheet sheet;

            _tables.Add(tableName);

            // インデックスシートを取得
            if (!_book.TryGetWorksheet(INDEXSHEETNAME, out sheet))
            {
                sheet = _book.AddWorksheet(INDEXSHEETNAME);

                // 見出し書き込み
                sheet.Cell(1, IndexsheetTableNameColumnIndex).Value = "テーブル名";
                sheet.Cell(1, IndexsheetTableCountColumnIndex).Value = "レコード数";

                sheet.Column(IndexsheetTableNameColumnIndex).Width = 35;
                sheet.Column(IndexsheetTableCountColumnIndex).Width = 15;
            }

            // 追記
            sheet.Cell(_tables.Count + 1, IndexsheetTableNameColumnIndex).Value = tableName;
            sheet.Cell(_tables.Count + 1, IndexsheetTableCountColumnIndex).Value = count;
        }

        private const string INDEXSHEETNAME = "index";
        private const int IndexsheetTableNameColumnIndex = 2;
        private const int IndexsheetTableCountColumnIndex = 3;
    }

    public class Progress
    {
        public static void Set(string message)
        {
            Console.Write(message);
        }

        public static void AppendProgress()
        {
            Console.Write(".");
        }

        public static void Clear()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }
    }
}
