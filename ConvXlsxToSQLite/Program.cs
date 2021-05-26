using System;
using GemBox.Spreadsheet;
using Microsoft.Data.Sqlite;
using System.IO;
using System.Linq;

namespace ConvertXlsxToSQLite
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter 1 to convert .xlsx to .db \n or \n enter 2 to upload images into db \n Your choise:");
            if (int.TryParse(Console.ReadLine(), out int ch))
            {
                if (ch == 1)
                {
                    SpreadsheetInfo.SetLicense("ERDC-HMNA-Q3HP-01TI");
                    ExcelFile workbook = ExcelFile.Load(@"D:\dbase\Verenich\baseNew.xlsx");
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
                }
                else if (ch == 2)
                {
                    FileInfo file;
                    string address, number, i11, i12, i21, i22;
                    string pathFolder = @"D:\dbase\Verenich\foto";
                    string dirName = string.Empty, imageName = string.Empty, imageSize = string.Empty;
                    byte[] imageData = new byte[0];
                    byte[] imageTemp = new byte[0];
                    if (Directory.Exists(pathFolder))
                    {
                        using (var connection = new SqliteConnection("Data Source = D:\\dbase\\Verenich\\tech.db"))
                        {
                            connection.Open();
                            SqliteCommand command;
                            /////////////////////////////////////////////
                            string[] directories = Directory.GetDirectories(pathFolder);
                            foreach (string subdir in directories)
                            {
                                dirName = subdir.Substring(subdir.LastIndexOf('\\') + 1);
                                string[] fields = dirName.Split('_');
                                address = fields[0]; number = fields[1]; i11 = fields[2]; i12 = fields[3]; i21 = fields[4]; i22 = fields[5];
                                //---
                                Console.WriteLine(dirName);
                                foreach (string str in fields)
                                    Console.WriteLine(str);
                                //---
                                imageName = string.Empty;
                                imageSize = string.Empty;
                                imageData = new byte[0];
                                imageTemp = new byte[0];
                                foreach (string str_file in Directory.GetFiles(subdir))
                                {
                                    file = new FileInfo(str_file);
                                    imageName += file.Name + ";";
                                    Console.WriteLine(imageName + " - ");
                                    using (FileStream fs = new FileStream(str_file, FileMode.Open))
                                    {
                                        imageTemp = new byte[fs.Length];
                                        imageSize += imageTemp.Length.ToString() + (";");
                                        fs.Read(imageTemp, 0, imageTemp.Length);
                                    }
                                    imageData = imageData.Concat(imageTemp).ToArray();
                                }
                                command = new SqliteCommand();
                                command.Connection = connection;
                                command.CommandText = @"UPDATE MALFUNCTIONS SET imageData = @ImageData, imageName = @ImageName, imageSize = @ImageSize " +
                                     " WHERE address = @Address AND number = @Number AND i11 = @I11 AND i12 = @I12 AND i21 = @I21 AND i22 = @I22";
                                command.Parameters.Add(new SqliteParameter("@ImageData", imageData));
                                command.Parameters.Add(new SqliteParameter("@ImageName", imageName));
                                command.Parameters.Add(new SqliteParameter("@ImageSize", imageSize));
                                command.Parameters.Add(new SqliteParameter("@Address", address));
                                command.Parameters.Add(new SqliteParameter("@Number", number));
                                command.Parameters.Add(new SqliteParameter("@I11", i11));
                                command.Parameters.Add(new SqliteParameter("@I12", i12));
                                command.Parameters.Add(new SqliteParameter("@I21", i21));
                                command.Parameters.Add(new SqliteParameter("@I22", i22));
                                int num = command.ExecuteNonQuery();
                                Console.WriteLine($"Added {num} objects.");
                            }
                            connection.Close();
                        }
                    }
                }
            }
            Console.ReadKey();
        }
    }
}
