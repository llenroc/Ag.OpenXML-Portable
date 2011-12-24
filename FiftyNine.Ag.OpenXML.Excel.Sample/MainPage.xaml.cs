using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using FiftyNine.Ag.OpenXML.Common.Storage;
using FiftyNine.Ag.OpenXML.Excel.Parts;

namespace FiftyNine.Ag.OpenXML.Excel.Sample
{
    public partial class MainPage : UserControl
    {
        public MainPage()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog();
            dlg.Filter = "Excel Document (.xlsx)|*.xlsx|Zip Files (.zip)|*.zip";
            dlg.DefaultExt = ".xlsx";
            if (dlg.ShowDialog() == true)
            {
                var doc = new SpreadsheetDocument();
                doc.ApplicationName = "SilverSpreadsheet";
                doc.Creator = "Chris Klug";
                doc.Company = "Intergen";

                var str1 = doc.Workbook.SharedStrings.AddString("Column 1");
                var str2 = doc.Workbook.SharedStrings.AddString("Column 2");
                var str3 = doc.Workbook.SharedStrings.AddString("Column 3");

                doc.Workbook.Sheets[0].Sheet.Rows[0].Cells[0].SetValue(str1);
                doc.Workbook.Sheets[0].Sheet.Rows[0].Cells[1].SetValue(str2);
                doc.Workbook.Sheets[0].Sheet.Rows[0].Cells[2].SetValue(str3);

                doc.Workbook.Sheets[0].Sheet.Rows[1].Cells[0].SetValue("Value 1");
                doc.Workbook.Sheets[0].Sheet.Rows[1].Cells[1].SetValue(1);
                doc.Workbook.Sheets[0].Sheet.Rows[1].Cells[2].SetValue(1001);

                doc.Workbook.Sheets[0].Sheet.Rows[2].Cells[0].SetValue("Value 2");
                doc.Workbook.Sheets[0].Sheet.Rows[2].Cells[1].SetValue(2);
                doc.Workbook.Sheets[0].Sheet.Rows[2].Cells[2].SetValue(1002);

                doc.Workbook.Sheets[0].Sheet.Rows[3].Cells[0].SetValue("Value 3");
                doc.Workbook.Sheets[0].Sheet.Rows[3].Cells[1].SetValue(3);
                doc.Workbook.Sheets[0].Sheet.Rows[3].Cells[2].SetValue(1003);

                doc.Workbook.Sheets[0].Sheet.Rows[4].Cells[0].SetValue("Value 4");
                doc.Workbook.Sheets[0].Sheet.Rows[4].Cells[1].SetValue(4);
                doc.Workbook.Sheets[0].Sheet.Rows[4].Cells[2].SetValue(1004);

                TablePart table = doc.Workbook.Sheets[0].Sheet.AddTable("My Table", "My Table", doc.Workbook.Sheets[0].Sheet.Rows[0].Cells[0], doc.Workbook.Sheets[0].Sheet.Rows[4].Cells[2]);
                table.TableColumns[0].Name = str1.String;
                table.TableColumns[1].Name = str2.String;
                table.TableColumns[2].Name = str3.String;

                doc.Workbook.Sheets[0].Sheet.AddColumnSizeDefinition(0, 2, 20);

                doc.Workbook.Sheets[0].Sheet.Rows[5].Cells[1].SetValue("Sum:");
                doc.Workbook.Sheets[0].Sheet.Rows[5].Cells[2].Formula = "SUM(" + doc.Workbook.Sheets[0].Sheet.Rows[1].Cells[2].CellName + ":" + doc.Workbook.Sheets[0].Sheet.Rows[4].Cells[2].CellName + ")";

                using (IStreamProvider storage = new ZipStreamProvider(dlg.OpenFile()))
                {
                    doc.Save(storage);
                }
            }
        }
    }
}
