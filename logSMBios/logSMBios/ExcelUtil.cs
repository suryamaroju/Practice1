using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using Practice;
using NPOI.HSSF.UserModel;

namespace logSMBios
{
    class ExcelUtil
    {
        public string OutPutFilePath = @"d:\SMBiosOutput.xls";
      // public  static Practice.Program p = new Practice.Program();
        public void CreateExcelFile()
        {
            try
            {
                using (FileStream stream = new FileStream(OutPutFilePath, FileMode.Create, FileAccess.Write))
                {
                    IWorkbook wb = new NPOI.HSSF.UserModel.HSSFWorkbook();
                    ISheet sheet = wb.CreateSheet("DIMMDetails");
                    ICreationHelper cH = wb.GetCreationHelper();

                    string[] Headers = { "Manufacturer", "AssetTag", "BankLocator", "DeviceLocator", "PartNumber", "SerialNumber", "ClockSpeed", "ExtendedSize", "Size" };
                    IRow row = sheet.CreateRow(0);
                    ICell cell;
                    for (int i = 0; i < Headers.Length; i++)
                    {
                        cell = row.CreateCell(i + 1);
                        cell.SetCellValue(Headers[i]);
                    }
                    for (int i = 1; i < 5; i++)
                    {
                        row = sheet.CreateRow(i);
                        for (int j = 0; j < 3; j++)
                        {
                            cell = row.CreateCell(j);

                        }
                    }
                    wb.Write(stream);
                }
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
            }
        }

        public void ExcelWrite(Details det, int i)
        {
            try
            {
                Details det1 = new Details();
                HSSFWorkbook hssfwb;
                IRow row;
                ICell cell;
                string sheetName = "DIMMDetails";

                using (FileStream file = new FileStream(OutPutFilePath, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new HSSFWorkbook(file);
                }
                ISheet sheet = hssfwb.GetSheet(sheetName) ?? hssfwb.CreateSheet(sheetName);


                row = sheet.GetRow(i) ?? sheet.CreateRow(i);
                cell = row.GetCell(0) ?? row.CreateCell(0);
                cell = row.GetCell(1) ?? row.CreateCell(1);
                cell.SetCellValue(det.manufacturer);
                cell = row.GetCell(2) ?? row.CreateCell(2);
                cell.SetCellValue(det.AssetTag);
                cell = row.GetCell(3) ?? row.CreateCell(3);
                cell.SetCellValue(det.BankLocator);
                cell = row.GetCell(4) ?? row.CreateCell(4);
                cell.SetCellValue(det.DeviceLocator);
                cell = row.GetCell(5) ?? row.CreateCell(5);
                cell.SetCellValue(det.PartNumber);
                cell = row.GetCell(6) ?? row.CreateCell(6);
                cell.SetCellValue(det.SerialNumber);
                cell = row.GetCell(7) ?? row.CreateCell(7);
                cell.SetCellValue(det.ClockSpeed);
                cell = row.GetCell(8) ?? row.CreateCell(8);
                cell.SetCellValue(det.ExtndSize);
                cell = row.GetCell(9) ?? row.CreateCell(9);
                cell.SetCellValue(det.Size);
                using (FileStream fs = new FileStream(OutPutFilePath, FileMode.Create, FileAccess.Write))
                {
                    hssfwb.Write(fs);
                }
            }catch(Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
            }
        }

        public void MergeCells(string Node, int startRow, int EndRow)
        {
            HSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(OutPutFilePath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new HSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet("DIMMDetails");

            IRow row;
            ICell cell;
            if (startRow != EndRow)
            {
                row = sheet.GetRow(startRow);
                cell = row.CreateCell(0);
                cell.SetCellValue(Node);
            }
           
            using (FileStream file = new FileStream(OutPutFilePath, FileMode.Open, FileAccess.Write))
            {
                hssfwb.Write(file);
            }

        }


    }
}
