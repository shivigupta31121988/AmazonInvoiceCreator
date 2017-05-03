using Microsoft.Office.Interop.Word;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;

namespace AmazonInvoice
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();

        }



        private void OpenFileDialogImplement()
        {

            var lDialog = new System.Windows.Forms.OpenFileDialog();

            lDialog.Title = "Open worksheet";
            lDialog.CheckFileExists = true;
            lDialog.Filter = ("Xls/Xlsx/Xlsm (Excel)|*.xls;*.xlsx;*.xlsm|Xls (Excel 2003)|_*.xls | Xlsx / Xlsm / Xlsb(Excel 2007) | *.xlsx; *.xlsm; *.xlsb | Xml(Xml) | _*.xml | Html(Html) | *.html | Csv(Csv) | *.csv | All files | *.* ");
            lDialog.InitialDirectory = @"D:\";
            lDialog.Multiselect = false;

            if (lDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                System.Windows.Forms.Application.DoEvents();
                try
                {
                    FileInfo File = new FileInfo(lDialog.FileName);
                    var fs = new FileStream(lDialog.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    var globalPackage = new ExcelPackage(File);
                    //  globalPackage.Load(fs);

                    var globalFileName = globalPackage.File.FullName;
                    var globalLastActiveWorksheet = globalPackage.Workbook.Worksheets.FirstOrDefault(f => f.View.TabSelected);
                    CreateDocument(globalLastActiveWorksheet);
                }
                finally
                {
                    lDialog.Dispose();
                }

            }

        }

        private System.Data.DataTable getTableFromWorksheet(ExcelWorksheet oSheet)
        {
            var dt = default(System.Data.DataTable);
            try
            {
                if ((oSheet.Dimension == null))
                {
                    dt = new System.Data.DataTable();
                }
                else
                {
                    int totalRows = oSheet.Dimension.End.Row;
                    int totalCols = oSheet.Dimension.End.Column;
                    dt = new System.Data.DataTable(oSheet.Name);
                    DataRow dr = null;
                    for (int i = 1; i <= totalCols; i++)
                    {
                        dt.Columns.Add(oSheet.Cells[1, i].Text);
                        //adding custom column names
                        //to display in form
                    }
                    for (int i = 1; i <= totalRows; i++)
                    {
                        dr = dt.Rows.Add();
                        for (int j = 1; j <= totalCols; j++)
                        {
                            dr[j - 1] = oSheet.Cells[i, j].Text;
                        }
                    }
                }

            }
            catch (Exception lException)
            {
                MessageBox.Show(lException.Message);
                return new System.Data.DataTable();
            }
            return dt;
        }

        //Create document method
        private void CreateDocument(ExcelWorksheet worksheet)
        {
            try
            {
                var dt = getTableFromWorksheet(worksheet);
                for (int i = 1; i < dt.Rows.Count; i++)
                {
                    //Create an instance for word app
                    Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();


                    winword.Visible = false;

                    //Create a missing variable for missing value
                    object missing = System.Reflection.Missing.Value;

                    //Create a new document
                    Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);


                    Microsoft.Office.Interop.Word.Paragraph paraHeading = document.Content.Paragraphs.Add(ref missing);
                    object styleHeading = "Heading 1";
                    paraHeading.Range.set_Style(ref styleHeading);
                    //  para1.Range.Text = "Para 1 text";
                    paraHeading.Range.Text = " \t\t\t\t\t\tINVOICE #0" + dt.Rows[i]["invoice-no"] + "\n\n";
                    //  paraHeading.Range.InsertAlignmentTab(WdParagraphAlignment.wdAlignParagraphCenter, 1);
                    //  paraHeading.Range.Font.Bold = 2;
                    //    paraHeading.Range.Font.Size = 18;
                    // paraHeading.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paraHeading.Range.InsertParagraphAfter();
                    //adding to address to document
                    document.Content.SetRange(0, 0);
                    document.Paragraphs.SpaceAfter = 0f;
                    document.Paragraphs.SpaceBefore = 0f;
                    //Microsoft.Office.Interop.Word.Paragraph paraRightInfo = document.Content.Paragraphs.Add(ref missing);
                    //object styleHeading1 = "Normal";
                    //paraRightInfo.Range.set_Style(ref styleHeading1);
                    ////  para1.Range.Text = "Para 1 text";
                    ////paraRightInfo.Range.Text = "INVOICE";
                    //paraRightInfo.Range.Text = "INVOICE #0" + worksheet.Cells[i, 30].Text;
                    //paraRightInfo.Range.Text +="DATE "+ DateTime.Now.Date.ToShortDateString();
                    ////  paraRightInfo.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    //paraRightInfo.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    //paraRightInfo.Range.InsertParagraphAfter();
                    //Add paragraph with Heading 1 style
                    Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add();
                    object styleHeading1 = "Normal";
                    para1.Range.set_Style(ref styleHeading1);
                    //  para1.Range.Text = "Para 1 text";
                    para1.Range.Text = "TO:\t\t\t\t\t\t\t\t\t\t" + "DATE " + DateTime.Now.Date.ToShortDateString();
                    // para1.Range.Font.Bold = 2;
                    para1.Range.Text += dt.Rows[i]["recipient-name"];
                    para1.Range.Text += dt.Rows[i]["ship-address-1"];
                    if (dt.Rows[i]["ship-address-2"].ToString() != "") para1.Range.Text += dt.Rows[i]["ship-address-2"];
                    if (dt.Rows[i]["ship-address-3"].ToString() != "") para1.Range.Text += dt.Rows[i]["ship-address-3"];
                    para1.Range.Text += dt.Rows[i]["ship-city"];
                    para1.Range.Text += dt.Rows[i]["ship-state"];
                    para1.Range.Text += dt.Rows[i]["ship-postal-code"];
                    para1.Range.Text += "Phone: " + dt.Rows[i]["buyer-phone-number"];
                    //  para1.Range.InsertParagraphBefore();
                    para1.Range.InsertParagraphAfter();

                    para1.SpaceAfter = 0f;
                    para1.SpaceBefore = 0f;
                    para1.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;

                    // para1.LineSpacing = 1f;

                    //Add paragraph with Heading 2 style
                    Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add();
                    object styleHeading2 = "Normal";
                    para2.Range.set_Style(ref styleHeading2);
                    //from address:
                    para2.Range.Text += "From:";
                    para2.Range.Text += "Aashi/Avi";
                    para2.Range.Text += "205, RKA , General Ganj";
                    para2.Range.Text += "Kanpur, Uttar Pradesh";
                    para2.Range.Text += "208001";
                    para2.Range.Text += "8097202529";
                    para2.Range.Text += "Tin No: 09337501255\n";
                    //  para2.SpaceAfterAuto=1;
                    para2.Range.InsertParagraphAfter();
                    para2.SpaceAfter = 0f;
                    para2.SpaceBefore = 0f;
                    para2.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;

                    //  para2.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;

                    //Create a 5X5 table and insert some dummy record


                    Table firstTable = document.Tables.Add(para1.Range, 5, 4, ref missing, ref missing);

                    //   firstTable.Range.Text
                    firstTable.Cell(1, 1).Range.Text = "Quantity";
                    //   firstTable.Cell(1, 1).Range.Font.Bold = 2;
                    firstTable.Cell(1, 2).Range.Text = "Description";
                    //  firstTable.Cell(1, 2).Range.Font.Bold = 2;
                    firstTable.Cell(1, 3).Range.Text = "Unit Price";
                    // firstTable.Cell(1, 3).Range.Font.Bold = 2;
                    firstTable.Cell(1, 4).Range.Text = "Total";
                    // firstTable.Cell(1, 4).Range.Font.Bold = 2;
                    int totalPrice = (Convert.ToInt16(dt.Rows[i]["quantity-purchased"]) * Convert.ToInt16(dt.Rows[i]["item-price"]));
                    firstTable.Cell(2, 1).Range.Text = dt.Rows[i]["quantity-purchased"].ToString();
                    firstTable.Cell(2, 2).Range.Text = dt.Rows[i]["product-name"].ToString();
                    firstTable.Cell(2, 3).Range.Text = "Rs." + dt.Rows[i]["item-price"].ToString();
                    firstTable.Cell(2, 4).Range.Text = "Rs." + totalPrice.ToString();

                    //  firstTable.Cell(3,1).Merge
                    firstTable.Cell(3, 3).Range.Text = "SUBTOTAL (Inclusive Tax)";
                    // firstTable.Cell(3, 1).Borders.Enable = 0;
                    //firstTable.Cell(3, 2).Borders.Enable = 0;
                    firstTable.Cell(3, 4).Range.Text = "Rs." + totalPrice.ToString();

                    firstTable.Cell(4, 3).Range.Text = "SHIPPING";
                    firstTable.Cell(4, 4).Range.Text = "Rs. 80";

                    firstTable.Cell(5, 3).Range.Text = "TOTAL";
                    firstTable.Cell(5, 4).Range.Text = "Rs." + Convert.ToInt16(totalPrice + 80);

                    firstTable.Borders.Enable = 1;
                    Microsoft.Office.Interop.Word.Cell cell;

                    firstTable.Rows[3].Cells[1].Merge(firstTable.Rows[3].Cells[2]);
                    firstTable.Rows[4].Cells[1].Merge(firstTable.Rows[4].Cells[2]);
                    firstTable.Rows[5].Cells[1].Merge(firstTable.Rows[5].Cells[2]);

                    //for (int colCounter = 1; colCounter <=3; colCounter++)
                    //{
                    //    cell = firstTable.Cell(colCounter,1);
                    //    cell.Merge(firstTable.Cell( colCounter,cell.RowIndex+1 ));
                    //}
                    //  firstTable.Rows[5].Cells[2].Merge(firstTable.Rows[4].Cells[1]);
                    // firstTable.Rows[4].Cells[1].Merge(firstTable.Rows[5].Cells[1]);
                    //  firstTable.Borders.Shadow = true;

                    //Add paragraph with normal  style  end letter
                    Microsoft.Office.Interop.Word.Paragraph paraLetter = document.Content.Paragraphs.Add();
                    //object styleHeadingLetter = "Normal";
                    paraLetter.Range.set_Style(ref styleHeading1);
                    //  paraLetter.Range.Text = "Para 1 text";
                    // paraLetter.Range.Text = "Dear Buyer,";
                    var strLetterRange = new StringBuilder();
                    strLetterRange.Append("\n\n\n");
                    strLetterRange.Append("Dear Buyer, \n");
                    strLetterRange.Append("Customer Satisfaction will always be our topmost priority. \n");
                    strLetterRange.Append("If you like our product and services please provide feedback.\n");
                    strLetterRange.Append("In case any issue with our product or services, kindly get back to us.\n");
                    strLetterRange.Append("We will be happy to resolve the issue and improve our services\n");
                    strLetterRange.Append("For Bulk order you can connect us directly @8097202529\n");
                    strLetterRange.Append("*** Thank You again for giving us a chance to serve you ***");
                    //paraLetter.Range.InsertParagraphAfter();
                    paraLetter.Range.Text = strLetterRange.ToString();
                 //    paraLetter.Borders.Shadow=true;
                    paraLetter.Range.InsertParagraphAfter();
                    paraLetter.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    paraLetter.SpaceAfter = 0f;
                    paraLetter.SpaceBefore = 0f;

                    //Save the document
                    object filename = @"D:\Invoice\\Invoice #" + dt.Rows[i]["invoice-no"].ToString();
                    document.SaveAs(ref filename);
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                    winword.Quit(ref missing, ref missing, ref missing);
                    winword = null;
                }
                MessageBox.Show("Documents created successfully !");
                //  this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            OpenFileDialogImplement();
        }
    }
}