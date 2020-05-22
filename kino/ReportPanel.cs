using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;


namespace ReportsGenerator
{
    public partial class ReportPanel : Form
    {
        public ReportPanel()
        {
            InitializeComponent();
            
        }
        private void GenerateButton_Click (object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedIndex == 0)
                    GenerateAllMoviesReport();
                else if (comboBox1.SelectedIndex == 3)
                    GenerateSalariesReport();
                else if (DateFrom.Value.Date > DateTo.Value.Date)
                {                    
                        MessageBox.Show("Date "+ DateFrom.Value.ToShortDateString() + " is older than date " + DateTo.Value.ToShortDateString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                }
                else if (comboBox1.SelectedIndex == 1)
                    GenerateWorkTimeReport();
                else if (comboBox1.SelectedIndex == 2)
                {
                    if (comboBox2.SelectedIndex == -1)
                    {
                        MessageBox.Show("Choose User ID", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    GenerateIndividualWorkTimeReport();
                }          
                else if (comboBox1.SelectedIndex == 4)
                {
                    if (comboBox2.SelectedIndex == -1)
                    {
                        MessageBox.Show("Choose User ID", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    GenerateIndividualSalaryReport();
                }
                else if (comboBox1.SelectedIndex == 5)
                    GenerateIncomesReport();
                else if (comboBox1.SelectedIndex == 6)
                    GenerateFoodSaleReport();
                else
                {
                    MessageBox.Show("Choose report to generate", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                                 
                    ExportDataTableToPdf(dataTable, "C:/Users/Public/Documents/" + comboBox1.Text + ".pdf", comboBox1.Text);
                    dataTable.Columns.Clear();
                    dataTable.Rows.Clear();
                    if (dataTable.Rows.Count > 0 && dataTable.Rows[0][0] != DBNull.Value)
                        MessageBox.Show("Generated report: " + comboBox1.Text + ".", "Report generated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                        MessageBox.Show("Generated empty report: " + comboBox1.Text + ".", "Report generated", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    if (checkBox1.Checked)
                    {
                        System.Diagnostics.Process.Start(@"C:/Users/Public/Documents/" + comboBox1.Text + ".pdf");
                        this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
                    }                                      
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Something went wrong");
            }    
            
        }

        private DataTable dataTable = new DataTable();

        private void GenerateAllMoviesReport()
        {
            using (SqlConnection conn = new SqlConnection(Helper.CnnString("kino")))
            {
                string procedure = "AllMoviesReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }
        }
        private void GenerateWorkTimeReport()
        {
            using (SqlConnection conn = new SqlConnection(Helper.CnnString("kino")))
            {
                string procedure = "WorkTimeReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.Parameters.AddWithValue("@DateFrom", DateFrom.Value);
                cmd.Parameters.AddWithValue("@DateTo", DateTo.Value);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }
        }
        private void GenerateIndividualWorkTimeReport()
        {
            
            using (SqlConnection conn = new SqlConnection(Helper.CnnString("kino")))
            {
                
                string procedure = "IndyvidualWorkTimeReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.Parameters.AddWithValue("@EmployeeId", comboBox2.Text);
                cmd.Parameters.AddWithValue("@DateFrom", DateFrom.Value);
                cmd.Parameters.AddWithValue("@DateTo", DateTo.Value);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }

        }
        private void GenerateSalariesReport()
        {
            using (SqlConnection conn = new SqlConnection(Helper.CnnString("kino")))
            {
                string procedure = "SalaryReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }
        }
        private void GenerateIndividualSalaryReport()
        {
            using (SqlConnection conn = new SqlConnection(Helper.CnnString("kino")))
            {
                string procedure = "IndyvidualSalaryReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.Parameters.AddWithValue("@EmployeeId", comboBox2.Text);
                cmd.Parameters.AddWithValue("@DateFrom", DateFrom.Value);
                cmd.Parameters.AddWithValue("@DateTo", DateTo.Value);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }
        }
        private void GenerateIncomesReport()
        {
            using (SqlConnection conn= new SqlConnection(Helper.CnnString("kino")))
            {
                string procedure = "IncomesReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.Parameters.AddWithValue("@DateFrom", DateFrom.Value);
                cmd.Parameters.AddWithValue("@DateTo", DateTo.Value);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }
        }
        private void GenerateFoodSaleReport()
        {
            using (SqlConnection conn = new SqlConnection(Helper.CnnString("kino")))
            {
                string procedure = "FoodSaleReport";
                SqlCommand cmd = new SqlCommand(procedure, conn);
                cmd.Parameters.AddWithValue("@DateFrom", DateFrom.Value);
                cmd.Parameters.AddWithValue("@DateTo", DateTo.Value);
                cmd.CommandType = CommandType.StoredProcedure;
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dataTable);
                da.Dispose();
            }
        }
        private void ExportDataTableToPdf(DataTable dtblTable, String strPdfPath, string strHeader)
        {

            FileStream fs = new FileStream(strPdfPath, FileMode.Create, FileAccess.Write, FileShare.None);
            Document document = new Document();
            document.SetPageSize(iTextSharp.text.PageSize.A4);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            //Header
            BaseFont bfntHead = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntHead = new Font(bfntHead, 16, 1, BaseColor.BLACK);
            Paragraph prgHeading = new Paragraph();
            prgHeading.Alignment = Element.ALIGN_CENTER;
            prgHeading.Add(new Chunk(strHeader.ToUpper(), fntHead));
            document.Add(prgHeading);
            
            document.Add(new Chunk("\n", fntHead));

            //Table
            PdfPTable table = new PdfPTable(dtblTable.Columns.Count);

            BaseFont btnColumnHeader = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            Font fntColumnHeader = new Font(btnColumnHeader, 10, 1, BaseColor.WHITE);
            for (int i = 0; i < dtblTable.Columns.Count; i++)
            {
                PdfPCell cell = new PdfPCell();
                cell.BackgroundColor = BaseColor.GRAY;
                cell.AddElement(new Chunk(dtblTable.Columns[i].ColumnName.ToUpper(), fntColumnHeader));
                table.AddCell(cell);
            }
            //Data
            for (int i = 0; i < dtblTable.Rows.Count; i++)
            {
                for (int j = 0; j < dtblTable.Columns.Count; j++)
                {
                    table.AddCell(dtblTable.Rows[i][j].ToString());
                }
            }

            document.Add(table);
            document.Close();
            writer.Close();
            fs.Close();
        }      

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }


    }
}
