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

using Excel = Microsoft.Office.Interop.Excel;
using Point = Microsoft.Office.Interop.PowerPoint;

namespace Principal
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application oApp;
            Excel.Workbook oBook;
            Excel.Worksheet oSheet;

            oApp = new Excel.Application();
            oBook = oApp.Workbooks.Add(Type.Missing);

            oSheet = (Excel.Worksheet)oBook.Worksheets.get_Item(1);

            oSheet.Name = "Principal";

            oSheet.Cells[1, 1] = "Some value";

            oBook.SaveAs(Directory.GetCurrentDirectory() + "\\ejemplo.xlsx");

            oBook.Close();

            oApp.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Point.Application oApp;
            Point.Presentation pptPresentation;

            Point.Slides oSlides;
            Point._Slide oSlide;
            Point.TextRange objText;

            //Create Presentation Text

            oApp = new Point.Application();

            pptPresentation = oApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

            Point.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Point.PpSlideLayout.ppLayoutText];

            //Create Slide
            oSlides = pptPresentation.Slides;

            oSlide = oSlides.AddSlide(1, customLayout);

            //Add Tittle
            objText = oSlide.Shapes[1].TextFrame.TextRange;
            objText.Text = "Prueba.com";
            objText.Font.Name = "Arial";
            objText.Font.Size = 28;

            //Body

            objText = oSlide.Shapes[2].TextFrame.TextRange;
            objText.Text = "One text\nTwo Text\nThree Text";

            oSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Es el primer ejemplo para la creación de presentaciones";
            pptPresentation.SaveAs(Directory.GetCurrentDirectory() + "\\presentacion.pptx", Point.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);

            pptPresentation.Close();

            oApp.Quit();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application oApp;
            Excel.Workbook oBook;
            

            oApp = new Excel.Application();

            oBook = oApp.Workbooks.Open(Directory.GetCurrentDirectory() + "ejemplo.xlsx");

            Excel._Worksheet sheet = oBook.Sheets[1];

            Excel.Range oRange = sheet.UsedRange;

            //Get Column and Rows
            int RowCount = oRange.Rows.Count;
            int ColumnCount = oRange.Columns.Count;

            //Get Value
            Valor.Text = oRange.Cells[1, 1].Value2.ToString();

            //Cleanup

            GC.Collect();
            GC.WaitForPendingFinalizers();

            oBook.Close();
            oApp.Quit();
            
        }
    }
}
