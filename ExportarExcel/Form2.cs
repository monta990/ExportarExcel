using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp;
using iTextSharp.text;
using System.IO;

namespace ExportarExcel
{
    public partial class Form2 : Form
    {
        Random R = new Random();

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < 10; i++)
            {
                dataGridView1.Rows.Add(
                    Guid.NewGuid().ToString().Substring(0, 10),
                    Guid.NewGuid().ToString().Substring(0, 12),
                    Guid.NewGuid().ToString().Substring(0, 14),
                    R.Next(1, 500) + "." + R.Next().ToString().Substring(0, 2));
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnPDF_Click(object sender, EventArgs e)
        {
            Document pdf = new Document(PageSize.LETTER.Rotate());
            try
            {
                PdfWriter.GetInstance(pdf, new FileStream("productos.pdf", FileMode.Create));
                pdf.Open();
                //inicio de la generación del pdf
                PdfPTable Tabla = new PdfPTable(5); //cantidad de columnas PDF
                PdfPCell Titulo = new PdfPCell(new Phrase("Titulo del Reporte"));
                Titulo.HorizontalAlignment=1; //1 para centrar
                Titulo.Colspan = 5;
                Tabla.AddCell(Titulo); //header pdf
                Tabla.AddCell("Titulo 1"); //titulo columna
                Tabla.AddCell("Titulo 2");
                Tabla.AddCell("Titulo 3");
                Tabla.AddCell("Titulo 4");
                Tabla.AddCell("Imagen");
                for (int i = 0; i < dataGridView1.Rows.Count; i++) //leer datagrid
                {
                    Tabla.AddCell(dataGridView1[0, i].Value.ToString());
                    Tabla.AddCell(dataGridView1[1, i].Value.ToString());
                    Tabla.AddCell(dataGridView1[2, i].Value.ToString());
                    Tabla.AddCell(dataGridView1[3, i].Value.ToString());
                    Tabla.AddCell(new PdfPCell(iTextSharp.text.Image.GetInstance(@"D:\miller64.png")));
                }
                pdf.Add(Tabla);
                //fin del pdf
                pdf.Close();
            }
            catch (DocumentException PDFerror)
            {

            }
            catch (IOException IOerror)
            {

            }
            System.Diagnostics.Process.Start("productos.pdf");
        }
    }
}
