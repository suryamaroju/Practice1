using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.Util;
using Practice;
using logSMBios;

namespace Practice
{
    public struct Details
    {
        public string manufacturer;
        public string SerialNumber;
        public string AssetTag;
        public string PartNumber;
        public string DeviceLocator;
        public string BankLocator;
        public string ClockSpeed;
        public string ExtndSize;
        public string Size;

    };
    class Program
    {     
        static void Main(string[] args)
        {

            Program p = new Program();
            ExcelUtil ex = new ExcelUtil();
            string line = "";
            int i = 1;
            int rowstart = 1, rowend = 1, outputexcelflag=0;
            string[] key = new string[2];
            List<string> Data = new List<string>();
            string path = "";

            string Manf = "";
        
            Data = p.UserInput();
            if (Data.Count == 2)
            {
                path = Data[0];

               Manf = Data[1];
            }
            int ManfDimmfound = 0, SmBIOSfileFound = 0;

            ex.CreateExcelFile();

            try
            {
                string manufacture = "";
                Details det = new Details();
                //string path = @"\\CSIQUAL01\CompQual\TestResults\Samsung\DIMM\M393A4K40CB2-CTD\Lenovo\TestPlan-140356\Dimm\TestPass-2018.09.07_19.34.55\TEC-00101-LogSmbios-140360";
                string[] nodefolders = Directory.GetDirectories(path);
               // List<string> data = new List<string>();

                foreach (string folder in nodefolders)
                {
                    string nodeName = Path.GetFileName(folder);
                    string[] files = Directory.GetFiles(folder);

                    foreach (string file in files)
                    {
                        if (file.Contains("SMBIOS-Type17"))
                        {
                            if (new FileInfo(file).Length != 0)

                            { 

                                SmBIOSfileFound = 1;
                            using (StreamReader reader = new StreamReader(file))
                            {
                                while ((line = reader.ReadLine()) != null)
                                {

                                    if (line.ToLower().Contains(Manf.ToLower()))
                                    {
                                        ManfDimmfound = 1;
                                        break;

                                    }
                                }
                                if (ManfDimmfound == 0)
                                {
                                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                                    Console.WriteLine("No DIMMS with the entered manufacturer found");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("Please re-execute the .EXE by providing the valid manufacturer");
                                    Console.ReadLine();
                                    System.Environment.Exit(1);
                                }

                            }



                            using (StreamReader reader = new StreamReader(file))
                            {
                                while ((line = reader.ReadLine()) != null)
                                {

                                    if (line.Contains("Device] (Type"))
                                    {

                                        while ((line = reader.ReadLine()) != null)
                                        {


                                            if (line.Contains("Device Locator"))
                                            {
                                                key = line.Split('"');
                                                det.DeviceLocator = key[1];
                                            }

                                            else if (line.Contains("Bank Locator"))
                                            {
                                                key = line.Split('"');
                                                det.BankLocator = key[1];
                                            }

                                            else if (line.Contains("Manufacturer"))
                                            {
                                                key = line.Split('"');
                                                manufacture = key[1];
                                                det.manufacturer = key[1];
                                            }

                                            else if (line.Contains("Serial Number"))
                                            {
                                                key = line.Split('"');
                                                det.SerialNumber = key[1];
                                            }

                                            else if (line.Contains("Asset Tag"))
                                            {
                                                key = line.Split('"');
                                                det.AssetTag = key[1];
                                            }

                                            else if (line.Contains("Part Number"))
                                            {
                                                key = line.Split('"');
                                                det.PartNumber = key[1];
                                            }
                                            else if (line.Contains("Clock Speed"))
                                            {
                                                key = line.Split('-');
                                                det.ClockSpeed = key[1];
                                            }

                                            else if (line.Contains("Extended Size"))
                                            {
                                                string[] ext = line.Split('-');
                                                det.ExtndSize = ext[ext.Length - 1].Trim();
                                            }
                                            else if (line.Contains("Size"))
                                            {
                                                string[] size = line.Split(',');
                                                det.Size = size[size.Length - 1].Trim();
                                            }

                                            if (det.manufacturer != null && det.SerialNumber != null && det.BankLocator != null
                                                 && det.AssetTag != null && det.DeviceLocator != null && det.PartNumber != null
                                                 && det.ClockSpeed != null && det.ExtndSize != null && det.Size != null)
                                            {
                                                if (manufacture.ToLower().Contains(Manf.ToLower()))
                                                {
                                                        Console.ForegroundColor = ConsoleColor.White;
                                                    Console.WriteLine(det.manufacturer);
                                                    Console.WriteLine(det.DeviceLocator);
                                                    Console.WriteLine(det.BankLocator);
                                                    Console.WriteLine(det.PartNumber);
                                                    Console.WriteLine(det.AssetTag);
                                                    Console.WriteLine(det.SerialNumber);

                                                    ex.ExcelWrite(det, i);
                                                    i++;
                                                    rowend++;
                                                }
                                                det.manufacturer = null;
                                                det.BankLocator = null;
                                                det.AssetTag = null;
                                                det.DeviceLocator = null;
                                                det.PartNumber = null;
                                                det.SerialNumber = null;
                                                det.ExtndSize = null;
                                                det.ClockSpeed = null;
                                                det.Size = null;

                                                break;
                                            }
                                        }
                                    }

                                }



                            }
                            ////////
                            if (manufacture == "")
                            {
                                Console.ForegroundColor = ConsoleColor.DarkRed;
                                Console.WriteLine("No DIMMS with the entered manufacturer found");
                                Console.ForegroundColor = ConsoleColor.White;
                                Console.WriteLine("Please enter correct manufacturer after re-executing the .EXE ");
                                 System.Environment.Exit(1);
                            }

                        }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine("File is empty: "+ Path.GetFileName(file)+ " in "+ Path.GetFileName(Path.GetDirectoryName(file)));
                                Console.ReadLine();
                                
                           
                            }

                        }
                        

                        ex.MergeCells(nodeName, rowstart, rowend);
                        rowstart = rowend;
                    }
                   if(SmBIOSfileFound == 0)
                        {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("File: SMBIOS-Type17 - Not found in " + nodeName);
                        Console.WriteLine();
                        Console.ReadLine();
                        
                    }
                    outputexcelflag = ManfDimmfound;
                    ManfDimmfound = 0;


                }
                if (outputexcelflag != 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("-------------------------------------------------------");
                    Console.WriteLine("Output file is created at D:\\SMBiosOutput.xls");
                    Console.WriteLine("-------------------------------------------------------");
                    Console.ReadLine();
                }
            }
            catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(e.Message);
                Console.WriteLine("Please enter the valid folder path");
                Console.ReadLine();
            }

        }

        public List<string> UserInput()
        {
            List<string> Data = new List<string>();
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Please enter path of the TEC-00101-LogSmbios path");
            string path = Console.ReadLine();
            DirectoryInfo dr = new DirectoryInfo(path);
            
            if((Path.GetFileName(path).ToLower().Contains("logsmbios"))  && (Directory.Exists(path)) )
                {
                string[] folders = Directory.GetDirectories(path);
                if (folders.Length != 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Please enter the manufacturer of DIMM");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("--------------------------------------------------------------------");
                    Console.WriteLine("For example: ");
                    Console.WriteLine("Samsung");
                    Console.WriteLine("Hynix");
                    Console.WriteLine("Micron");
                    Console.WriteLine("--------------------------------------------------------------------");
                    string Manf = Console.ReadLine().Trim();
                    Data.Add(path);
                    Data.Add(Manf);
                    return Data;
                }
                return Data;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Please provide the correct path of LOGSmbios folder by executing the application again");
                Console.WriteLine(" Check whether the files are present the folders.");
                            
                Console.WriteLine("--------------------------------------------------------------------");
                Console.WriteLine("Process exited");
                Console.WriteLine("--------------------------------------------------------------------");             
                Console.ReadLine();
                Console.ReadLine();
                Console.ReadLine();
                Console.ReadLine();
                return Data;
            }
           
        }
/*
 
        public void ExcelWrite(Details det, int i)
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
        }

        public void CreateExcelFile()
        {
            using (FileStream stream = new FileStream(OutPutFilePath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook wb = new HSSFWorkbook();
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
            //CellRangeAddress cra = new CellRangeAddress(startRow, EndRow, 0,0);
            //sheet.AddMergedRegion(cra);
            using (FileStream file = new FileStream(OutPutFilePath, FileMode.Open, FileAccess.Write))
            {
                hssfwb.Write(file);
            }

        }
        */


    }
}


