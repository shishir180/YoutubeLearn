using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;
using System.Text;

namespace WebApplication1
{
    public partial class About : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //getDataTable();
            StringBuilder sc = new StringBuilder();
            sc.Append("Hii");
            sc.Append("Bye");
            Console.WriteLine(sc);
        }
        public void getDataTable()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Name");
            dt.Columns.Add("Age");
            dt.Columns.Add("Salary");
            dt.Columns.Add("Gender");
            dt.AcceptChanges();
            dt.Rows.Add("My Name is Shishir Singh\njibesh mishra is His Name Hi father Name is Gandu Mishra\nsagar sahu is friend of jini", "89456543", "544550\n8776656565", "Male");
            dt.Copy();
            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    int size = Convert.ToString(dt.Rows[i]["Name"]).Length;
            //    if (size > 15)
            //    {

            //        dt.Rows[i]["Name"] = Convert.ToString(dt.Rows[i]["Name"]).Replace("\n","")+"\n";
            //    }
            //}

            //DatatableToExcel(dt);
            //DatatableToExcel3(dt);
            getData7(dt);
            StringBuilder sc= new StringBuilder();
            sc.Append("Hii");
            sc.Append("Bye");
            Console.WriteLine(sc.ToString());
        }
        //private void DatatableToExcel(DataTable dt)
        //{
        //    //using (XLWorkbook wb = new XLWorkbook())
        //    //{
        //    //    wb.Worksheets.Add(dt, "Customers");

        //    //    Response.Clear();
        //    //    Response.Buffer = true;
        //    //    Response.Charset = "";
        //    //    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //    //    Response.AddHeader("content-disposition", "attachment;filename=CampaignReport.xlsx");
        //    //    using (MemoryStream MyMemoryStream = new MemoryStream())
        //    //    {
        //    //        wb.SaveAs(MyMemoryStream);
        //    //        MyMemoryStream.WriteTo(Response.OutputStream);
        //    //        Response.Flush();
        //    //        Response.End();
        //    //    }
        //    //}

        //    string attachment = "attachment; filename=TemplateReport.xls";
        //    HttpContext.Current.Response.ClearContent();
        //    HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        //    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());
        //    string tab = "";

        //    foreach (DataColumn dc in dt.Columns)
        //    {
        //        HttpContext.Current.Response.Write(tab + dc.ColumnName);
        //        tab = "\t";
        //    }
        //    HttpContext.Current.Response.Write("\n");
        //    int i;
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        tab = "";
        //        for (i = 0; i < dt.Columns.Count; i++)
        //        {
        //            string leadingChar = i == 0 ? "'" : "";
        //            HttpContext.Current.Response.Write(tab + leadingChar + dr[i].ToString().Replace('\n', ' ').Replace(Convert.ToChar(10), ' ').Replace(Convert.ToChar(13), ' '));
        //            tab = "\t";
        //        }
        //        HttpContext.Current.Response.Write("\n");
        //    }
        //    // Response.End();
        //    HttpContext.Current.Response.Flush();
        //    HttpContext.Current.Response.SuppressContent = true;
        //    HttpContext.Current.ApplicationInstance.CompleteRequest();
        //}

        //private void DatatableToExcel(DataTable dt)
        //{
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        var worksheet = wb.Worksheets.Add(dt, "Customers");

        //        // Loop through each cell in the worksheet and set text wrapping
        //        foreach (var cell in worksheet.CellsUsed())
        //        {
        //            // Set the text of the cell
        //            cell.Value = cell.Value;

        //            // Set text wrapping
        //            cell.Style.Alignment.WrapText = true;
        //        }

        //        Response.Clear();
        //        Response.Buffer = true;
        //        Response.Charset = "";
        //        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        Response.AddHeader("content-disposition", "attachment;filename=CampaignReport.xlsx");

        //        using (MemoryStream MyMemoryStream = new MemoryStream())
        //        {
        //            wb.SaveAs(MyMemoryStream);
        //            MyMemoryStream.WriteTo(Response.OutputStream);
        //            Response.Flush();
        //            Response.End();
        //        }
        //    }
        //}

        //private void DatatableToExcel(DataTable dt)
        //{
        //    string attachment = "attachment; filename=TemplateReport.xls";
        //    HttpContext.Current.Response.ClearContent();
        //    HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        //    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

        //    // Add the HTML style for text wrapping
        //    HttpContext.Current.Response.Write("<style>td {mso-number-format:\"\\@\";white-space:nowrap;text-align:left;}</style>");

        //    string tab = "";

        //    foreach (DataColumn dc in dt.Columns)
        //    {
        //        HttpContext.Current.Response.Write(tab + dc.ColumnName);
        //        tab = "\t";
        //    }
        //    HttpContext.Current.Response.Write("\n");
        //    int i;
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        tab = "";
        //        for (i = 0; i < dt.Columns.Count; i++)
        //        {
        //            string leadingChar = i == 0 ? "'" : "";
        //            HttpContext.Current.Response.Write(tab + leadingChar + dr[i].ToString().Replace('\n', ' ').Replace(Convert.ToChar(10), ' ').Replace(Convert.ToChar(13), ' '));
        //            tab = "\t";
        //        }
        //        HttpContext.Current.Response.Write("\n");
        //    }

        //    HttpContext.Current.Response.Flush();
        //    HttpContext.Current.Response.SuppressContent = true;
        //    HttpContext.Current.ApplicationInstance.CompleteRequest();
        //}
        //private void DatatableToExcel(DataTable dt)
        //{
        //    string attachment = "attachment; filename=TemplateReport.xls";
        //    HttpContext.Current.Response.ClearContent();
        //    HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        //    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

        //    // Write Excel header
        //    HttpContext.Current.Response.Write("<?xml version=\"1.0\"?>\n");
        //    HttpContext.Current.Response.Write("<?mso-application progid=\"Excel.Sheet\"?>\n");
        //    HttpContext.Current.Response.Write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

        //    // Write Excel styles
        //    HttpContext.Current.Response.Write("<Styles>\n");
        //    HttpContext.Current.Response.Write("<Style ss:ID=\"s1\">\n");
        //    HttpContext.Current.Response.Write("<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\n");
        //    HttpContext.Current.Response.Write("</Style>\n");
        //    HttpContext.Current.Response.Write("</Styles>\n");

        //    // Write worksheet
        //    HttpContext.Current.Response.Write("<Worksheet ss:Name=\"Sheet1\">\n");
        //    HttpContext.Current.Response.Write("<Table>\n");

        //    // Write column headers
        //    HttpContext.Current.Response.Write("<Row>\n");
        //    foreach (DataColumn dc in dt.Columns)
        //    {
        //        HttpContext.Current.Response.Write("<Cell><Data ss:Type=\"String\">" + dc.ColumnName + "</Data></Cell>\n");
        //    }
        //    HttpContext.Current.Response.Write("</Row>\n");

        //    // Write data rows
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        HttpContext.Current.Response.Write("<Row>\n");
        //        foreach (object o in dr.ItemArray)
        //        {
        //            HttpContext.Current.Response.Write("<Cell><Data ss:Type=\"String\">" + o.ToString() + "</Data></Cell>\n");
        //        }
        //        HttpContext.Current.Response.Write("</Row>\n");
        //    }

        //    // Close worksheet and workbook
        //    HttpContext.Current.Response.Write("</Table>\n");
        //    HttpContext.Current.Response.Write("</Worksheet>\n");
        //    HttpContext.Current.Response.Write("</Workbook>\n");

        //    HttpContext.Current.Response.Flush();
        //    HttpContext.Current.Response.SuppressContent = true;
        //    HttpContext.Current.ApplicationInstance.CompleteRequest();
        //}
        //private void DatatableToExcel(DataTable dt)
        //{
        //    string attachment = "attachment; filename=TemplateReport.xls";
        //    HttpContext.Current.Response.ClearContent();
        //    HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        //    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

        //    // Write Excel header
        //    HttpContext.Current.Response.Write("<?xml version=\"1.0\"?>\n");
        //    HttpContext.Current.Response.Write("<?mso-application progid=\"Excel.Sheet\"?>\n");
        //    HttpContext.Current.Response.Write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

        //    // Write Excel styles
        //    HttpContext.Current.Response.Write("<Styles>\n");
        //    HttpContext.Current.Response.Write("<Style ss:ID=\"s1\">\n");
        //    HttpContext.Current.Response.Write("<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\n");
        //    HttpContext.Current.Response.Write("</Style>\n");
        //    HttpContext.Current.Response.Write("</Styles>\n");

        //    // Write worksheet
        //    HttpContext.Current.Response.Write("<Worksheet ss:Name=\"Sheet1\">\n");
        //    HttpContext.Current.Response.Write("<Table>\n");

        //    // Write column headers
        //    HttpContext.Current.Response.Write("<Row>\n");
        //    foreach (DataColumn dc in dt.Columns)
        //    {
        //        HttpContext.Current.Response.Write("<Cell ss:StyleID=\"s1\"><Data ss:Type=\"String\">" + dc.ColumnName + "</Data></Cell>\n");
        //    }
        //    HttpContext.Current.Response.Write("</Row>\n");

        //    // Write data rows
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        HttpContext.Current.Response.Write("<Row>\n");
        //        foreach (object o in dr.ItemArray)
        //        {
        //            HttpContext.Current.Response.Write("<Cell ss:StyleID=\"s1\"><Data ss:Type=\"String\">" + o.ToString() + "</Data></Cell>\n");
        //        }
        //        HttpContext.Current.Response.Write("</Row>\n");
        //    }

        //    // Close worksheet and workbook
        //    HttpContext.Current.Response.Write("</Table>\n");
        //    HttpContext.Current.Response.Write("</Worksheet>\n");
        //    HttpContext.Current.Response.Write("</Workbook>\n");

        //    HttpContext.Current.Response.Flush();
        //    HttpContext.Current.Response.SuppressContent = true;
        //    HttpContext.Current.ApplicationInstance.CompleteRequest();
        //}

        //public void getData1(DataTable dt)
        //{
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        wb.Worksheets.Add(dt, "Customers");

        //        Response.Clear();
        //        Response.Buffer = true;
        //        Response.Charset = "";
        //        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        Response.AddHeader("content-disposition", "attachment;filename=CampaignReport.xlsx");
        //        using (MemoryStream MyMemoryStream = new MemoryStream())
        //        {
        //            wb.SaveAs(MyMemoryStream);
        //            MyMemoryStream.WriteTo(Response.OutputStream);
        //            Response.Flush();
        //            Response.End();
        //        }
        //    }
        //}
        //private void DatatableToExcel1(DataTable dt)
        //{
        //    string attachment = "attachment; filename=TemplateReport.xls";
        //    HttpContext.Current.Response.ClearContent();
        //    HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        //    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

        //    // Write Excel header
        //    HttpContext.Current.Response.Write("<?xml version=\"1.0\"?>\n");
        //    HttpContext.Current.Response.Write("<?mso-application progid=\"Excel.Sheet\"?>\n");
        //    HttpContext.Current.Response.Write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

        //    // Write Excel styles
        //    HttpContext.Current.Response.Write("<Styles>\n");
        //    HttpContext.Current.Response.Write("<Style ss:ID=\"s1\">\n");
        //    HttpContext.Current.Response.Write("<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\n");
        //    HttpContext.Current.Response.Write("</Style>\n");
        //    HttpContext.Current.Response.Write("</Styles>\n");

        //    // Write worksheet
        //    HttpContext.Current.Response.Write("<Worksheet ss:Name=\"Sheet1\">\n");
        //    HttpContext.Current.Response.Write("<Table>\n");

        //    // Write column headers
        //    HttpContext.Current.Response.Write("<Row>\n");
        //    foreach (DataColumn dc in dt.Columns)
        //    {
        //        HttpContext.Current.Response.Write("<Cell ss:StyleID=\"s1\"><Data ss:Type=\"String\">" + dc.ColumnName + "</Data></Cell>\n");
        //    }
        //    HttpContext.Current.Response.Write("</Row>\n");

        //    // Write data rows
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        HttpContext.Current.Response.Write("<Row>\n");
        //        foreach (object o in dr.ItemArray)
        //        {
        //            string cellValue = o.ToString().Replace("\n", "<br />");
        //            HttpContext.Current.Response.Write("<Cell ss:StyleID=\"s1\"><Data ss:Type=\"String\">" + cellValue + "</Data></Cell>\n");
        //        }
        //        HttpContext.Current.Response.Write("</Row>\n");
        //    }

        //    // Close worksheet and workbook
        //    HttpContext.Current.Response.Write("</Table>\n");
        //    HttpContext.Current.Response.Write("</Worksheet>\n");
        //    HttpContext.Current.Response.Write("</Workbook>\n");

        //    HttpContext.Current.Response.Flush();
        //    HttpContext.Current.Response.SuppressContent = true;
        //    HttpContext.Current.ApplicationInstance.CompleteRequest();
        //}
        //private void DatatableToExcel2(DataTable dt)
        //{
        //    string attachment = "attachment; filename=TemplateReport.xls";
        //    HttpContext.Current.Response.ClearContent();
        //    HttpContext.Current.Response.AddHeader("content-disposition", attachment);
        //    HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
        //    Response.ContentEncoding = System.Text.Encoding.Unicode;
        //    Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());

        //    // Write Excel header
        //    HttpContext.Current.Response.Write("<?xml version=\"1.0\"?>\n");
        //    HttpContext.Current.Response.Write("<?mso-application progid=\"Excel.Sheet\"?>\n");
        //    HttpContext.Current.Response.Write("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"\n");
        //    HttpContext.Current.Response.Write(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");

        //    // Write Excel styles
        //    HttpContext.Current.Response.Write("<Styles>\n");
        //    HttpContext.Current.Response.Write("<Style ss:ID=\"s1\">\n");
        //    HttpContext.Current.Response.Write("<Alignment ss:Vertical=\"Bottom\" ss:WrapText=\"1\"/>\n");
        //    HttpContext.Current.Response.Write("</Style>\n");
        //    HttpContext.Current.Response.Write("</Styles>\n");

        //    // Write worksheet
        //    HttpContext.Current.Response.Write("<Worksheet ss:Name=\"Sheet1\">\n");
        //    HttpContext.Current.Response.Write("<Table>\n");

        //    // Write column headers
        //    HttpContext.Current.Response.Write("<Row>\n");
        //    foreach (DataColumn dc in dt.Columns)
        //    {
        //        HttpContext.Current.Response.Write("<Cell ss:StyleID=\"s1\"><Data ss:Type=\"String\">" + dc.ColumnName + "</Data></Cell>\n");
        //    }
        //    HttpContext.Current.Response.Write("</Row>\n");

        //    // Write data rows
        //    foreach (DataRow dr in dt.Rows)
        //    {
        //        HttpContext.Current.Response.Write("<Row>\n");
        //        foreach (object o in dr.ItemArray)
        //        {
        //            string cellValue = o.ToString().Replace("\n", "<br />\n");
        //            HttpContext.Current.Response.Write("<Cell ss:StyleID=\"s1\"><Data ss:Type=\"String\">" + cellValue + "</Data></Cell>\n");
        //        }
        //        HttpContext.Current.Response.Write("</Row>\n");
        //    }

        //    // Close worksheet and workbook
        //    HttpContext.Current.Response.Write("</Table>\n");
        //    HttpContext.Current.Response.Write("</Worksheet>\n");
        //    HttpContext.Current.Response.Write("</Workbook>\n");

        //    HttpContext.Current.Response.Flush();
        //    HttpContext.Current.Response.SuppressContent = true;
        //    HttpContext.Current.ApplicationInstance.CompleteRequest();
        //}
        //private void DatatableToExcel3(DataTable dt)
        //{
        //    using (var package = new ExcelPackage())
        //    {
        //        // Add a worksheet to the package
        //        var worksheet = package.Workbook.Worksheets.Add("Sheet1");

        //        // Loop through the rows and columns of the DataTable and add data to the worksheet
        //        for (int rowIndex = 1; rowIndex <= dt.Rows.Count; rowIndex++)
        //        {
        //            for (int colIndex = 1; colIndex <= dt.Columns.Count; colIndex++)
        //            {
        //                var cellValue = dt.Rows[rowIndex - 1][colIndex - 1].ToString();

        //                // If the column contains '\n', split it and add to multiple rows
        //                if (cellValue.Contains("\n"))
        //                {
        //                    var splitValues = cellValue.Split('\n');
        //                    for (int i = 0; i < splitValues.Length; i++)
        //                    {
        //                        worksheet.Cells[rowIndex + i, colIndex].Value = splitValues[i];
        //                    }
        //                }
        //                else
        //                {
        //                    worksheet.Cells[rowIndex, colIndex].Value = cellValue;
        //                }
        //            }
        //        }

        //        // Save the Excel package to a memory stream
        //        using (MemoryStream stream = new MemoryStream())
        //        {
        //            package.SaveAs(stream);

        //            // Return the Excel file as a byte array
        //            byte[] excelBytes = stream.ToArray();

        //            // You can now provide a way for the user to download the Excel file, for example, in an ASP.NET MVC Controller, you can use the FileResult:
        //            // return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "yourfilename.xlsx");
        //        }
        //    }
        //}
        //public void getData(DataTable dt)
        //{
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        var worksheet = wb.Worksheets.Add("Customers");

        //        for (int rowIndex = 1; rowIndex <= dt.Rows.Count; rowIndex++)
        //        {
        //            for (int colIndex = 1; colIndex <= dt.Columns.Count; colIndex++)
        //            {
        //                var cellValue = dt.Rows[rowIndex - 1][colIndex - 1].ToString();

        //                // Replace newline characters with Excel newline and add to a single cell
        //                cellValue = cellValue.Replace("\n", Environment.NewLine);

        //                worksheet.Cell(rowIndex, colIndex).Value = cellValue;
        //            }
        //        }

        //        Response.Clear();
        //        Response.Buffer = true;
        //        Response.Charset = "";
        //        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        Response.AddHeader("content-disposition", "attachment;filename=CampaignReport.xlsx");
        //        using (MemoryStream MyMemoryStream = new MemoryStream())
        //        {
        //            wb.SaveAs(MyMemoryStream);
        //            MyMemoryStream.WriteTo(Response.OutputStream);
        //            Response.Flush();
        //            Response.End();
        //        }
        //    }
        //}

        //public void getData4(DataTable dt)
        //{
        //    using (XLWorkbook wb = new XLWorkbook())
        //    {
        //        var worksheet = wb.Worksheets.Add("Customers");

        //        for (int rowIndex = 1; rowIndex <= dt.Rows.Count; rowIndex++)
        //        {
        //            for (int colIndex = 1; colIndex <= dt.Columns.Count; colIndex++)
        //            {
        //                var cellValue = dt.Rows[rowIndex - 1][colIndex - 1].ToString();

        //                // Replace newline characters with Excel newline and add to a single cell
        //                cellValue = cellValue.Replace("\n", Environment.NewLine);

        //                worksheet.Cell(rowIndex, colIndex).Value = cellValue;
        //            }
        //        }

        //        // AutoFit the columns based on content
        //        worksheet.Columns().AdjustToContents();

        //        Response.Clear();
        //        Response.Buffer = true;
        //        Response.Charset = "";
        //        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        //        Response.AddHeader("content-disposition", "attachment;filename=CampaignReport.xlsx");
        //        using (MemoryStream MyMemoryStream = new MemoryStream())
        //        {
        //            wb.SaveAs(MyMemoryStream);
        //            MyMemoryStream.WriteTo(Response.OutputStream);
        //            Response.Flush();
        //            Response.End();
        //        }
        //    }
        //}
        public void getData7(DataTable dt)
        {
            using (XLWorkbook wb = new XLWorkbook())
            {
                var worksheet = wb.Worksheets.Add("Customers");

                // Add header row
                for (int colIndex = 1; colIndex <= dt.Columns.Count; colIndex++)
                {
                    worksheet.Cell(1, colIndex).Value = dt.Columns[colIndex - 1].ColumnName;
                }

                // Populate data rows
                for (int rowIndex = 2; rowIndex <= dt.Rows.Count + 1; rowIndex++)
                {
                    for (int colIndex = 1; colIndex <= dt.Columns.Count; colIndex++)
                    {
                        var cellValue = dt.Rows[rowIndex - 2][colIndex - 1].ToString();

                        // Replace newline characters with Excel newline and add to a single cell
                        cellValue = cellValue.Replace("\n", Environment.NewLine);

                        worksheet.Cell(rowIndex, colIndex).Value = cellValue;
                    }
                }

                // AutoFit the columns based on content
                worksheet.Columns().AdjustToContents();

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=CampaignReport.xlsx");
                using (MemoryStream MyMemoryStream = new MemoryStream())
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();
                    Response.End();
                }
            }
        }

    }
}