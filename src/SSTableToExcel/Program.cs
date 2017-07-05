using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SSTableToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("PC Name:");
            var pcName = Console.ReadLine();
            Console.Write("Database Name:");
            var dbName = Console.ReadLine();
            Console.Write("Login Name:");
            var userName = Console.ReadLine();
            Console.Write("Password:");
            var password = Console.ReadLine();
            
            var setting = new Models.DBSetting() { サーバー名 = pcName, データベース名 = dbName, ユーザー名 = userName, パスワード = password };

            try
            {
                var db = new Models.DBAccess();
                db.Getter(setting);

                Console.WriteLine("保存終了しました。");
            }
            catch (Exception e)
            {
                // 例外はすべてさらけ出す
                Console.WriteLine(e);
            }

            Console.WriteLine("なにかキーを押してください");
            Console.ReadKey();
        }
    }
}
