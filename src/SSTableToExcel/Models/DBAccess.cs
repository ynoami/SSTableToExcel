using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data.SqlClient;
using Dapper;
using Newtonsoft.Json;
using System.IO;

namespace SSTableToExcel.Models
{
    public class DBAccess
    {
        private SqlConnection GetConnection(DBSetting setting)
        {
            var connectionString = $"Server=tcp:{setting.サーバー名};Database={setting.データベース名};User ID={setting.ユーザー名};Password={setting.パスワード};Trusted_Connection=False;Connection Timeout=30;";
            var con = new SqlConnection(connectionString);
            con.Open();
            return con;
        }

        public void Getter(DBSetting setting)
        {
            using (var con = GetConnection(setting))
            {
                // 先頭が「M_」から始まるテーブル名を取得
                var names = con.Query<Dictionary<string, object>>("select name from sysobjects where xtype = 'U' and name like 'M_%' order by name", null).Select(_ => _.Values.First() as string).ToArray();

                // 出力先ディレクトリ作成(雑)
                var directoryName = $"{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                Directory.CreateDirectory(directoryName);
                Directory.CreateDirectory(Path.Combine(directoryName, "JSON"));
                Directory.CreateDirectory(Path.Combine(directoryName, "Excel"));
                
                var book = new SSTableExcelBook();
                foreach (var tableName in names)
                {
                    // 全フィールド及びフィールド名を取得(dynamicのまま)[インジェクション対策はしていません！]
                    var records = con.Query<Dictionary<string, object>>($"select * from {tableName}", null).ToList();
                    var columnNames = con.Query<Dictionary<string, object>>($"select COLUMN_NAME from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '{tableName}' order by ORDINAL_POSITION", null)
                        .Select(_ => _.Values.First() as string).ToArray();

                    //// JSON形式で出力
                    WriteJson(records, directoryName, tableName);

                    //// Excelにシートを追加
                    book.AppendExcelSheet(records, tableName, columnNames);
                }

                // Excel bookを保存
                book.SaveAs($"{directoryName}/Excel/All.xlsx");
            }
        }

        private void WriteJson(List<Dictionary<string, object>> records, string directoryName, string tableName)
        {
            var serializer = new JsonSerializer();
            using (var writer = new StreamWriter($"{directoryName}/JSON/{tableName}.json"))
            {
                serializer.Serialize(writer, records);
            }
        }
    }
}
