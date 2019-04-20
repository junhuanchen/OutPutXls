using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace PutXls
{
    class Program
    {
        struct XlsData
        {
            public string data;
            public string name;
            public string test;
        }

        const string PathName = @"../DATA/";

        static void OutPutXls(string ExcelName, XlsData Data, DateTime Now)
        {
            new Thread(r =>
            {
                lock (PathName)
                {

                RE:
                    try
                    {
                        HSSFWorkbook hssfworkbook = new HSSFWorkbook(new FileStream(PathName + ExcelName, FileMode.Open, FileAccess.Read));

                        ISheet sheet = hssfworkbook.GetSheet("record");

                        if (sheet.LastRowNum > 0)
                        {
                            sheet.ShiftRows(sheet.FirstRowNum + 1, sheet.LastRowNum, 1);

                            if (sheet.LastRowNum > 5000)
                            {
                                System.IO.File.Move(PathName + ExcelName, PathName + ExcelName.Insert(ExcelName.Length - 4, Now.ToString("-yyyy MM dd HH mm ss")));
                                // sheet.RemoveRow(sheet.GetRow(5000));
                            }
                        }

                        // ADD
                        var cells = sheet.CreateRow(sheet.FirstRowNum + 1);
                        cells.CreateCell(0).SetCellValue(Now.ToString());
                        var fields = Data.GetType().GetFields();
                        for (int i = 0; i < fields.Length; i++)
                        {
                            var cell = cells.CreateCell(i + 1);
                            cell.CellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
                            cell.SetCellValue(fields[i].GetValue(Data).ToString());
                        }

                        FileStream file = new FileStream(PathName + ExcelName, FileMode.Open, FileAccess.ReadWrite, FileShare.Write);
                        hssfworkbook.Write(file);
                        file.Close();
                    }
                    catch (FileNotFoundException ex)
                    {
                        try
                        {
                            HSSFWorkbook hssfworkbook = new HSSFWorkbook();
                            ISheet sheet = hssfworkbook.CreateSheet("record");
                            // INIT
                            var top_cells = sheet.CreateRow(0);
                            top_cells.CreateCell(0).SetCellValue("record time");
                            sheet.AutoSizeColumn(0);
                            var fields = Data.GetType().GetFields();
                            for (int i = 0; i < fields.Length; i++)
                            {
                                top_cells.CreateCell(i + 1).SetCellValue("   " + fields[i].Name + "   ");
                                sheet.AutoSizeColumn(i);
                            }
                            FileStream file = new FileStream(PathName + ExcelName, FileMode.OpenOrCreate);
                            hssfworkbook.Write(file);
                            file.Close();
                            goto RE;
                        }
                        catch (Exception e)
                        {
                            ;
                        }
                    }
                    catch (Exception ex)
                    {
                        ;
                    }
                }
            }).Start();
        }

        static void Main(string[] args)
        {
            if (!System.IO.Directory.Exists("../DATA"))
            {
                System.IO.Directory.CreateDirectory("../DATA");
            }
            var now = DateTime.Now.ToString("yyyy-MM-dd HH mm ss");

            OutPutXls(string.Format("test1-{0}.xls", now), new XlsData() { data = "123", name = "456", test = "789" }, DateTime.Now);
            OutPutXls(string.Format("test2-{0}.xls", now), new XlsData() { data = "abc", name = "def", test = "ghj" }, DateTime.Now);
            OutPutXls(string.Format("test2-{0}.xls", now), new XlsData() { data = "ghj", name = "def", test = "abc" }, DateTime.Now);
            OutPutXls(string.Format("test1-{0}.xls", now), new XlsData() { data = "789", name = "456", test = "123" }, DateTime.Now);
            
        }
    }
}
