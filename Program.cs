using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

namespace nordXtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            var inputFileLocation = args[0];
            var outputFileLocation = args[1];
            var nordAccounts = new List<NordAccount>();

            string rawNordVPNAccounts = System.IO.File.ReadAllText(inputFileLocation);
            Regex emailRegex = new Regex(@"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*:[a-zA-Z0-9]+",
            RegexOptions.IgnoreCase);
            MatchCollection emailMatches = emailRegex.Matches(rawNordVPNAccounts);
            StringBuilder sb = new StringBuilder();
            foreach (Match emailMatch in emailMatches)
            {
                sb.AppendLine(emailMatch.Value);
                var rawAccount = emailMatch.Value.Split(':');
                nordAccounts.Add(new NordAccount(rawAccount[0], rawAccount[1]));
            }
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(nordAccounts), (typeof(DataTable)));
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputFileLocation, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }
    }

    class NordAccount
    {
        public string username;
        public string password;
        public string accountStatus;
        public string vpnStatus;

        public NordAccount(string username, string password, string accountStatus = "NOTCHECKED", string vpnStatus = "NOTCHECKED")
        {
            this.username = username;
            this.password = password;
            this.accountStatus = accountStatus;
            this.vpnStatus = vpnStatus;
        }

    }
}
