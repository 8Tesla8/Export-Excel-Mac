using System;
using System.IO;
using AppKit;
using Foundation;
using OfficeOpenXml;

namespace ExcelMac2 {
    public partial class ViewController : NSViewController {
        public ViewController(IntPtr handle) : base(handle) { }

        private const string XlsxContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        public override void ViewDidLoad() {
            base.ViewDidLoad();

            // Do any additional setup after loading the view.


            var fileName = ShowSaveFileWindow("test", "/Users/Martin/Desktop/");

            var newFile = new FileInfo(fileName);
            if (newFile.Exists) {
                Console.WriteLine("File exists");

                newFile.Delete();
                using (FileStream fs = File.Create(fileName)) { }
            }
            else {
                Console.WriteLine("File not exists");
                using (FileStream fs = File.Create(fileName)) { }
            }

            using (var package = CreateExcelPackage()) {
                package.SaveAs(new FileInfo(fileName));
            }
        }


        public string ShowSaveFileWindow(string name, string path) {
            var saveWindow = NSSavePanel.SavePanel;
            saveWindow.AllowsOtherFileTypes = false;
            saveWindow.ExtensionHidden = true;
            saveWindow.CanCreateDirectories = true;
            saveWindow.CanSelectHiddenExtension = false;

            saveWindow.AllowedFileTypes = new string[] {
                "xlsx"
            };

            saveWindow.NameFieldStringValue = name;

            var tf = NSTextField.CreateLabel("xlsx");
            saveWindow.AccessoryView = tf;

            if (!string.IsNullOrEmpty(path))
                saveWindow.Directory = path;

            if (saveWindow.RunModal() == Convert.ToInt32(NSModalResponse.OK))
                return saveWindow.Url.Path;

            return null;
        }

        private ExcelPackage CreateExcelPackage() {
            var package = new ExcelPackage();
            package.Workbook.Properties.Title = "Salary Report";
            package.Workbook.Properties.Author = "Vahid N.";
            package.Workbook.Properties.Subject = "Salary Report";
            package.Workbook.Properties.Keywords = "Salary";


            var worksheet = package.Workbook.Worksheets.Add("Employee");

            //First add the headers
            worksheet.Cells[1, 1].Value = "ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "Gender";
            worksheet.Cells[1, 4].Value = "Salary (in $)";

            //Add values

            var numberformat = "#,##0";
            //var dataCellStyleName = "TableNumber";
            //var numStyle = package.Workbook.Styles.CreateNamedStyle(dataCellStyleName);
            //numStyle.Style.Numberformat.Format = numberformat;

            worksheet.Cells[2, 1].Value = 1000;
            worksheet.Cells[2, 2].Value = "Jon";
            worksheet.Cells[2, 3].Value = "M";
            worksheet.Cells[2, 4].Value = 5000;
            worksheet.Cells[2, 4].Style.Numberformat.Format = numberformat;

            worksheet.Cells[3, 1].Value = 1001;
            worksheet.Cells[3, 2].Value = "Graham";
            worksheet.Cells[3, 3].Value = "M";
            worksheet.Cells[3, 4].Value = 10000;
            worksheet.Cells[3, 4].Style.Numberformat.Format = numberformat;

            worksheet.Cells[4, 1].Value = 1002;
            worksheet.Cells[4, 2].Value = "Jenny";
            worksheet.Cells[4, 3].Value = "F";
            worksheet.Cells[4, 4].Value = 5000;
            worksheet.Cells[4, 4].Style.Numberformat.Format = numberformat;

            // Add to table / Add summary row
            var tbl = worksheet.Tables.Add(new ExcelAddressBase(fromRow: 1, fromCol: 1, toRow: 4, toColumn: 4), "Data");
            tbl.ShowHeader = true;
            //tbl.TableStyle = TableStyles.Dark9;
            tbl.ShowTotal = true;
            //tbl.Columns[3].DataCellStyleName = dataCellStyleName;
            //tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            worksheet.Cells[5, 4].Style.Numberformat.Format = numberformat;

            // AutoFitColumns
            //worksheet.Cells[1, 1, 4, 4].AutoFitColumns();

            //worksheet.HeaderFooter.OddFooter.InsertPicture(
            //new FileInfo(Path.Combine(_hostingEnvironment.WebRootPath, "images", "captcha.jpg")),
            //PictureAlignment.Right);

            return package;
        }

        public override NSObject RepresentedObject {
            get {
                return base.RepresentedObject;
            }
            set {
                base.RepresentedObject = value;
                // Update the view, if already loaded.
            }
        }
    }
}
