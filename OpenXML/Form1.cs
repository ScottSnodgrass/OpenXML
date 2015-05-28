using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXML
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fileName = @".\excel\test.xlsx";

            // OpenXMLドキュメントを読み込み
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // WorkbookPartを取得
                WorkbookPart wbPart = document.WorkbookPart;

                // Sheetを取得
                Sheet targetSheet = wbPart.Workbook.Descendants<Sheet>().Where(s => s.SheetId == 1).FirstOrDefault();
                if (targetSheet == null)
                {
                    return;
                }

                // WorksheetPartの取得
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(targetSheet.Id));

                // A1のCellオブジェクトを取得
                Cell cell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == "A1").FirstOrDefault();

                // cellの調査
                if (cell == null) return;
                if (cell.DataType.Value != CellValues.SharedString) return;

                // SharedStringItemの取得
                int targetIndex = int.Parse(cell.InnerText);
                SharedStringItem item = wbPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(targetIndex);

                //OUTPUT
                listBox1.Items.Add(item.InnerText);
            }



        }
    }
}
