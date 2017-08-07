using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportarExcel
{
    public partial class Form1 : Form
    {
        Random R = new Random();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < 10; i++)
            {
                dataGridView1.Rows.Add(
                    Guid.NewGuid().ToString().Substring(0,10),
                    Guid.NewGuid().ToString().Substring(0,12),
                    Guid.NewGuid().ToString().Substring(0,14),
                    R.Next(1,500)+"."+R.Next().ToString().Substring(0,2));
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            SaveFileDialog File = new SaveFileDialog();
            File.Filter = "Excel (*.xlsx)|*.xlsx";
            if (File.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application app; //selecciona la app
                Microsoft.Office.Interop.Excel.Workbook libro; //genera libro para excel
                Microsoft.Office.Interop.Excel.Worksheet hoja; //genera la hoja del libro
                app = new Microsoft.Office.Interop.Excel.Application();
                libro = app.Workbooks.Add();
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range Rango;
                for (int i = 1; i <=dataGridView1.Columns.Count; i++)
                {
                    hoja.Cells[1,i]=dataGridView1.Columns[i-1].HeaderText.ToString();
                    Rango = hoja.Cells[1,i];
                    Rango.Font.Bold = true;
                    Rango.Font.Name = "Comics Sans";
                    //Rango.Font.Size = 18;
                    Rango.Interior.Color = Color.Gray;
                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        hoja.Cells[2 + i, j + 1] = dataGridView1[j ,i].Value.ToString();
                    }
                }
                libro.SaveAs(File.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                libro.Close(true);
                app.Quit();
            }
        }
    }
}
