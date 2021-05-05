using System;
using GemBox.Spreadsheet;
using Microsoft.Data.Sqlite;

namespace ConvertXlsxToSQLite
{
    class Program
    {
        static void Main(string[] args)
        {
            // ограниечения бесплатной версии: 5 листов и 150 строк
            //SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            SpreadsheetInfo.SetLicense("ERDC-HMNA-Q3HP-01TI");
            ExcelFile workbook = ExcelFile.Load(@"D:\dbase\Verenich\base150.xlsx");
            ExcelWorksheet sheet = workbook.Worksheets[0];
            int i, address, number, i11, i12, i21, i22;
            Console.WriteLine("");
            string cause, purpose, method, consequences;
            using (var connection = new SqliteConnection("Data Source = D:\\dbase\\Verenich\\tech.db"))
            {
                connection.Open();
                SqliteCommand command = new SqliteCommand();
                command.Connection = connection;
                command.CommandText = "DELETE FROM MALFUNCTIONS";
                int num = command.ExecuteNonQuery();
                for (i = 0; i < sheet.Rows.Count; i++)
                {
                    address = sheet.Cells[i, 0].Value == null ? 0 : Convert.ToInt32(sheet.Cells[i, 0].StringValue);
                    number = sheet.Cells[i, 1].Value == null ? 0 : Convert.ToInt32(sheet.Cells[i, 1].StringValue);
                    i11 = sheet.Cells[i, 2].Value == null ? 0 : Convert.ToInt32(sheet.Cells[i, 2].StringValue);
                    i12 = sheet.Cells[i, 3].Value == null ? 0 : Convert.ToInt32(sheet.Cells[i, 3].StringValue);
                    i21 = sheet.Cells[i, 4].Value == null ? 0 : Convert.ToInt32(sheet.Cells[i, 4].StringValue);
                    i22 = sheet.Cells[i, 5].Value == null ? 0 : Convert.ToInt32(sheet.Cells[i, 5].StringValue);
                    cause = sheet.Cells[i, 6].Value == null ? "" : sheet.Cells[i, 6].StringValue;
                    purpose = sheet.Cells[i, 7].Value == null ? "" : sheet.Cells[i, 7].StringValue;
                    method = sheet.Cells[i, 8].Value == null ? "" : sheet.Cells[i, 8].StringValue;
                    consequences = sheet.Cells[i, 9].Value == null ? "" : sheet.Cells[i, 9].StringValue;
                    Console.WriteLine($"{address} {number} {i11} {i12} {i21} {i22} {cause} {purpose} {method} {consequences}");
                    command.CommandText = "INSERT INTO MALFUNCTIONS (address, number, i11, i12, i21, i22, cause, purpose, method, consequences) VALUES " +
                        $"({address}, {number}, {i11}, {i12}, {i21}, {i22}, '{cause}', '{purpose}', '{method}', '{consequences}')";
                    num = command.ExecuteNonQuery();
                }
                connection.Close();
            }
            Console.WriteLine("----------------------------------");
            Console.WriteLine($"Конвертировано {i} строк");

            Console.ReadKey();
            /*command.CommandText = "select * from MALFUNCTIONS";
            using (SqliteDataReader reader = command.ExecuteReader())
            {
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        address = Convert.ToInt32(reader.GetValue(1));
                        number = Convert.ToInt32(reader.GetValue(2));
                        i11 = Convert.ToInt32(reader.GetValue(3));
                        i12 = Convert.ToInt32(reader.GetValue(4));
                        i21 = Convert.ToInt32(reader.GetValue(5));
                        i22 = Convert.ToInt32(reader.GetValue(6));
                        cause = reader.GetValue(7).ToString();
                        purpose = reader.GetValue(8).ToString();
                        method = reader.GetValue(9).ToString();
                        consequences = reader.GetValue(10).ToString();
                        Console.WriteLine($"{address} {number} {i11} {i12} {i21} {i22} {cause} {purpose} {method} {consequences}");
                    }
                }
            }*/
        }
    }
}
