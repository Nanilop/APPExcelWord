using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word= Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelWordC
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            string ruta,dato;
            SaveFileDialog dialogoGuardar = new SaveFileDialog();
            dialogoGuardar.Filter = "Archivos Word (*.docx)|*.docx|Hoja de cálculo de Microsoft Excel (*.xlsx)|*.xlsx";
            if(dialogoGuardar.ShowDialog() != DialogResult.OK)
                        {
                return;
            }
            else if(dialogoGuardar.FilterIndex == 1){
                ruta= dialogoGuardar.FileName;
                var WordApp = new Word.Application();
                WordApp.Visible= true;
                WordApp.Documents.Add();
                dato = txtDato.Text;
                WordApp.Selection.TypeText(dato);
                WordApp.ActiveDocument.SaveAs2(ruta);
                MessageBox.Show("Archivo creado");
            }
            else if (dialogoGuardar.FilterIndex == 2)
            {

                Excel.Application aplicacion;
                Excel.Workbook libros_trabajo;
                Excel.Worksheet hoja_trabajo;
                aplicacion = new Excel.Application();
                aplicacion.Visible = true;
                libros_trabajo = aplicacion.Workbooks.Add();
                hoja_trabajo = libros_trabajo.Worksheets.get_Item(1);
                hoja_trabajo.Cells[1, 1] = txtDato.Text;
                libros_trabajo.SaveAs(dialogoGuardar.FileName);
                MessageBox.Show("Archivo creado");
            }

        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }
    }
}
