using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using ClosedXML.Excel;


namespace Golfzon_TVNX_Test.Class
{
    internal class Excel
    {
        LogControl LOG = new LogControl();

        public DataTable ExcelToDataTable(string filePath, string extension)
        {
            string conStr = "";
            switch (extension)
            {
                case ".xls":
                    conStr = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + filePath + "; Extended Properties =\"Excel 8.0;HDR={1};IMEX=0\"";
                    break;

                case ".xlsx":
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=0\"";
                    break;

                case ".xlsm":
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=0\"";
                    break;
            }

            OleDbConnection conn = new OleDbConnection(conStr);
            OleDbCommand cmd = new OleDbCommand();

            DataTable dt = new DataTable();

            conn.Open();
            DataTable dtExcelSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            // 첫시트의 이름을 받아옴
            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            OleDbDataAdapter oda = new OleDbDataAdapter(cmd.CommandText = "SELECT * FROM [" + SheetName + "]", conn);
            oda.Fill(dt);
            conn.Close();

            return dt;
        }


        // ALM 익스포트 파일 읽기 전용
        public DataTable ALMExcelToDataTable(string filePath, string extension)
        {
            string conStr = "";
            switch (extension)
            {
                case ".xls":
                    conStr = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = " + filePath + "; Extended Properties =\"Excel 8.0;HDR={1};IMEX=0\"";
                    break;

                case ".xlsx":
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=0\"";
                    break;

                case ".xlsm":
                    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR={1};IMEX=0\"";
                    break;
            }

            OleDbConnection conn = new OleDbConnection(conStr);
            OleDbCommand cmd = new OleDbCommand();
            DataTable dt = new DataTable();
            conn.Open();
            DataTable dtExcelSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            // 첫시트의 이름을 받아옴
            string SheetName = "Export$";
            OleDbDataAdapter oda = new OleDbDataAdapter(cmd.CommandText = "SELECT * FROM [" + SheetName + "]", conn);
            oda.Fill(dt);
            conn.Close();

            return dt;
        }


        // xlsx 형식 엑셀 익스포트
        public void Excel2007Export(DataTable dt, string fileName)
        {
            LOG.Write("엑셀 익스포트 시작");
            XLWorkbook wb = new XLWorkbook();
            wb.Worksheets.Add(dt, "CC 테스트 결과");            
            wb.SaveAs(fileName);
            LOG.Write("엑셀 익스포트 완료");         
        }


        public DataTable ExcelToDataTable2(string filePath)
        {

            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dt;
            }
        }
    }
}
