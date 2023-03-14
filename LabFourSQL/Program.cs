using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LabFourSQL
{
    class Program
    {
        static void Main(string[] args)
        {
            /*    string workFile = "C:\\#\\InputData.xlsm";
                List<Company> inputData = ReadDataExcelFileDOM(workFile);

             //   DBConnector.WriteDataDB(inputData);
            */

            List<Company> inputData = DBConnector.ReadDataFromDB();

            // Accounts Receivable Turnover == 46
            // Operating Gross Margin == 4
            // Cash/Current Liability == 59

            int RowNum, WNum = 0, MNum = 0;
            Dictionary<string, double> var46 = new Dictionary<string, double>();
            Dictionary<string, double> var4 = new Dictionary<string, double>();
            Dictionary<string, double> var59 = new Dictionary<string, double>();

            var46.Add("TotalM", 0);
            var46.Add("TotalW", 0);

            var4.Add("TotalM", 0);
            var4.Add("TotalW", 0);

            var59.Add("TotalM", 0);
            var59.Add("TotalW", 0);

            foreach (Company company in inputData)
            {
                if (company.getBankrupt())
                {
                    var46["TotalM"] += company.getAccountsReceivableTurnover();
                    var4["TotalM"] += company.getOperatingGrossMargin();
                    var59["TotalM"] += company.getCashCurrentLiability();
                    MNum++;
                }
                else
                {
                    var46["TotalW"] += company.getAccountsReceivableTurnover();
                    var4["TotalW"] += company.getOperatingGrossMargin();
                    var59["TotalW"] += company.getCashCurrentLiability();
                    WNum++;
                }
            }
            RowNum = WNum + MNum;

            countValuesOne(var46, MNum, WNum, RowNum);
            countValuesOne(var4, MNum, WNum, RowNum);
            countValuesOne(var59, MNum, WNum, RowNum);

            foreach (Company company in inputData)
            {
                if (company.getBankrupt())
                {   //SSm += (cell.value - AverM)^2
                    var46["SSm"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["AverM"]), 2);
                    var4["SSm"] += Math.Pow((company.getOperatingGrossMargin() - var4["AverM"]), 2);
                    var59["SSm"] += Math.Pow((company.getCashCurrentLiability() - var59["AverM"]), 2);

                    //SS += (cell.value - Average)^2
                    var46["SS"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["Average"]), 2);
                    var4["SS"] += Math.Pow((company.getOperatingGrossMargin() - var4["Average"]), 2);
                    var59["SS"] += Math.Pow((company.getCashCurrentLiability() - var59["Average"]), 2);
                }
                else
                {
                    //SSw += (cell.value - AverW)^2
                    var46["SSw"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["AverW"]), 2);
                    var4["SSw"] += Math.Pow((company.getOperatingGrossMargin() - var4["AverW"]), 2);
                    var59["SSw"] += Math.Pow((company.getCashCurrentLiability() - var59["AverW"]), 2);

                    //SS += (cell.value - Average)^2
                    var46["SS"] += Math.Pow((company.getAccountsReceivableTurnover() - var46["Average"]), 2);
                    var4["SS"] += Math.Pow((company.getOperatingGrossMargin() - var4["Average"]), 2);
                    var59["SS"] += Math.Pow((company.getCashCurrentLiability() - var59["Average"]), 2);
                }
            }

            countValuesTwo(var46, MNum, WNum);
            countValuesTwo(var4, MNum, WNum);
            countValuesTwo(var59, MNum, WNum);

            Console.WriteLine("Показатель 'Accounts Receivable Turnover'");
            WriteData(var46);
            Console.WriteLine("Показатель 'Operating Gross Margin'");
            WriteData(var4);
            Console.WriteLine("Показатель 'Cash/Current Liability'");
            WriteData(var59);
            Console.ReadLine();
        }
        private static void countValuesOne(Dictionary<string, double> variable, int MNum, int WNum, int RowNum)
        {

            variable.Add("AverM", variable["TotalM"] / MNum);                    //AverM = TotalM / MNum
            variable.Add("AverW", variable["TotalW"] / WNum);                   //AverW = TotalW / WNum
            variable.Add("Total", variable["TotalW"] + variable["TotalM"]);   //Total = TotalW + TotalM
            variable.Add("Average", variable["Total"] / RowNum);                //Average = Total / RowNum

            variable.Add("SSw", 0);
            variable.Add("SSm", 0);
            variable.Add("SS", 0);
        }

        private static void countValuesTwo(Dictionary<string, double> variable, int MNum, int WNum)
        {
            variable.Add("SSmist", variable["SSw"] + variable["SSm"]);                             //SSmist = SSw + SSm
            variable.Add("SSeff", WNum * Math.Pow(variable["AverW"] - variable["Average"], 2) +        //SSeff = WNum * (AverW - Average) ^ 2 + _
                                            MNum * Math.Pow(variable["AverM"] - variable["Average"], 2));     //   MNum * (AverM - Average) ^ 2
            variable.Add("D", variable["SSeff"] / variable["SS"] * 100);                           // D = SSeff / SS * 100

            //Для графика
            variable.Add("dolyaSSmist", variable["SSmist"] / variable["SS"] * 100);
            variable.Add("dolyaSSw", variable["SSw"] / variable["SS"] * 100);
            variable.Add("dolyaSSm", variable["SSm"] / variable["SS"] * 100);

        }

        private static void WriteData(Dictionary<string, double> variable)
        {
            Console.WriteLine("Влияние показателя на выходную переменную: " + Convert.ToString(Math.Round(variable["D"], 2)) + "%" +
                "\t\t\tНеобъясненная SS: " + Convert.ToString(Math.Round(variable["dolyaSSmist"], 2)) + "%");
            Console.WriteLine("Общая сумма квадратов отклонений: " + Convert.ToString(Math.Round(variable["SS"], 2)) +
                "\t\t\tДоля банкротов в общей ошибке: " + Convert.ToString(Math.Round(variable["dolyaSSm"], 2)) + "%");
            Console.WriteLine("Объясненная влиянием 'а' сум.кв.откл: " + Convert.ToString(Math.Round(variable["SSeff"], 2)) +
                "\t\t\tДоля не банкротов в общей ошибке: " + Convert.ToString(Math.Round(variable["dolyaSSw"], 2)) + "%");
            Console.WriteLine("Необъясненная сумма квадратов отклонений: " + Convert.ToString(Math.Round(variable["SSmist"], 2)));
            Console.WriteLine("\n");
        }
        //Считывание входных данных
        static List<Company> ReadDataExcelFileDOM(string fileName)
        {
            List<Company> output = new List<Company>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;       //Инициализация переменной книги 
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();      //Выбор первой части в книге
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();    //Выбор первого листа в книге

                int i = 0;
                bool headerFlag = true;     //Флаг заголовка

                foreach (Row row in sheetData.Elements<Row>())
                {
                    if (headerFlag)
                    { //пропускаем заголовок
                        headerFlag = false;
                        continue;
                    }
                    output.Add(new Company(false, 0, 0, 0));

                    //сохранение данных со столбцов строки во временные переменные 
                    List<Cell> rowData = row.Elements<Cell>().ToList();
                    string row1 = rowData[0].CellValue.Text;
                    string row46 = rowData[46].CellValue.Text;
                    string row4 = rowData[4].CellValue.Text;
                    string row59 = rowData[59].CellValue.Text;

                    if (row1 == "1")
                        output[i].setBankrupt(true);
                    else
                        output[i].setBankrupt(false);

                    if (string.IsNullOrEmpty(row46))
                        output[i].setAccountsReceivableTurnover(0);
                    else
                        output[i].setAccountsReceivableTurnover(Convert.ToDouble(row46.Replace('.', ',')));

                    if (string.IsNullOrEmpty(row4))
                        output[i].setOperatingGrossMargin(0);
                    else
                        output[i].setOperatingGrossMargin(Convert.ToDouble(row4.Replace('.', ',')));

                    if (string.IsNullOrEmpty(row59))
                        output[i].setCashCurrentLiability(0);
                    else
                        output[i].setCashCurrentLiability(Convert.ToDouble(row59.Replace('.', ',')));
                    i++;
                }
            }
            return output;
        }

        
    }
}
