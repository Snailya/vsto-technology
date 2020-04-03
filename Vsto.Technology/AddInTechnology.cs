using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Vsto.Technology
{
    public partial class AddInTechnology
    {
        private void AddInTechnology_Startup(object sender, EventArgs e)
        {
            // add a new handler on sheet before right click
            Globals.AddInTechnology.Application.SheetBeforeRightClick += ApplicationOnSheetBeforeRightClick;

            // add a new control
            _generateSheetsButton = (Office.CommandBarButton) Globals.AddInTechnology.Application
                .CommandBars["cell"].Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    missing, missing, missing, true);
            _generateSheetsButton.BeginGroup = true;
            _generateSheetsButton.Caption = "生成表";
            _generateSheetsButton.Click += _generateSheetsButton_Click;
            _generateSheetsButton.Move(Before: 1);

            _generateChildrenSheetsButton = (Office.CommandBarButton) Globals.AddInTechnology.Application
                .CommandBars["cell"].Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    missing, missing, missing, true);
            _generateChildrenSheetsButton.BeginGroup = true;
            _generateChildrenSheetsButton.Caption = "生成子表";
            _generateChildrenSheetsButton.Click += _generateChildrenSheetsButton_Click;
            _generateChildrenSheetsButton.Move(Before: 2);
        }


        private void AddInTechnology_Shutdown(object sender, EventArgs e)
        {
            // remove the control
            // XXX: better check existence before delete
            _generateSheetsButton?.Delete();
            _generateChildrenSheetsButton?.Delete();

            // remove sheet before right click handler
            Globals.AddInTechnology.Application.SheetBeforeRightClick -= ApplicationOnSheetBeforeRightClick;
        }

        #region VSTO generated code

        /// <summary>
        ///     Required method for Designer support - do not modify
        ///     the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += AddInTechnology_Startup;
            Shutdown += AddInTechnology_Shutdown;
        }

        #endregion

        #region Callbacks

        private void ApplicationOnSheetBeforeRightClick(object sh, Range target, ref bool cancel)
        {
            // store target
            _target = target;
        }

        /// <summary>
        ///     Generates quotation sheets by user selection using templates
        /// </summary>
        /// <param name="ctrl"></param>
        /// <param name="canceldefault"></param>
        private void _generateSheetsButton_Click(Office.CommandBarButton ctrl, ref bool canceldefault)
        {
            // find all group label by finding where column B is blank but column C is not
            var labels = ((Range) _target.Columns["B"].Cells).Cast<Range>().Select(cell => cell.Value2 as string)
                .Distinct().Where(value => value != null);
            // create a diction to store the starting address of each label
            var sectionRowLookup = ((Range) _target.Columns["B"].Cells).Cast<Range>()
                .ToLookup(c => c.Value2 as string, c => c.Row - _target[1, 1].Row as int?);

            // add a new workbook
            var customTemplatePath = ConfigurationManager.AppSettings["TemplatePath"];
            var templatePath = string.IsNullOrWhiteSpace(customTemplatePath)
                ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\自定义 Office 模板\涂装项目报价表模板.xltx"
                : customTemplatePath;
            var wb = _target.Application.Workbooks.Add(templatePath);

            // paste to new workbook
            foreach (Worksheet sheet in wb.Sheets)
            {
                // format summary style
                sheet.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;

                // continue if it's a constant sheet
                // todo: separate into app config
                if (sheet.Name == "材料及加工费综合单价") continue;

                // apply auto fill
                var startCell = sheet.UsedRange.End[XlDirection.xlDown].Offset[1, 0];
                var endCell = sheet.Cells[sheet.UsedRange.Rows.Count, sheet.UsedRange.Columns.Count]
                    .Offset[_target.Rows.Count - 1, 0] as Range;
                var source = startCell.Resize[ColumnSize: sheet.UsedRange.Columns.Count];
                var destination = sheet.Range[startCell.Address, endCell.Address];
                source.AutoFill(destination);

                // modify section sum
                foreach (var sectionRowGroup in sectionRowLookup)
                {
                    // convert to absolute row number of section
                    var absoluteRowGroup = sectionRowGroup.Select(i => i + startCell.Row);

                    if (string.IsNullOrEmpty(sectionRowGroup.Key))
                    {
                        // highlight section header
                        absoluteRowGroup.ToList().ForEach(row =>
                        {
                            var rng = (Range) sheet.Cells[row, 1].Resize[1, sheet.UsedRange.Columns.Count];
                            rng.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent5;
                            rng.Interior.TintAndShade = 0.5;
                        });
                        continue;
                    }

                    // write summary
                    var sumFormulaSource = (Range) sheet.Cells[absoluteRowGroup.Min() - 1, 1];
                    sumFormulaSource.Formula = $"=SUM(A{absoluteRowGroup.Min()}:A{absoluteRowGroup.Max()})";

                    var sumDestination = sheet.Range[sumFormulaSource.Address,
                        sheet.Cells[sumFormulaSource.Row, sheet.UsedRange.Columns.Count]];

                    sumFormulaSource.AutoFill(sheet.Range[sumFormulaSource.Address, sumDestination.Address],
                        XlAutoFillType.xlFillValues);

                    ((Range) sheet.Rows[$"{absoluteRowGroup.Min()}:{absoluteRowGroup.Max()}"]).Group();
                }

                // paste content
                // todo: separate into app config
                if (sheet.Name == "加工单价表")
                    _target.Resize[ColumnSize: 4].Copy();
                else
                    _target.Resize[ColumnSize: 3].Copy();
                sheet.Activate();
                startCell.PasteSpecial(XlPasteType.xlPasteValues);

                // modify coefficient if it's summary sheet
                if (sheet.Name == "报价表")
                {
                    _target.Columns[5].Copy();
                    endCell.Offset[-(_target.Rows.Count - 1), -2].Resize[_target.Rows.Count, 2]
                        .PasteSpecial(XlPasteType.xlPasteValues);
                }

                // adjust row height and column width
                sheet.UsedRange.Columns.AutoFit();
                sheet.UsedRange.Rows.AutoFit();
            }
        }

        /// <summary>
        ///     Generates children sheets by user selection and link to origin sheet.
        /// </summary>
        /// <param name="Ctrl"></param>
        /// <param name="CancelDefault"></param>
        private void _generateChildrenSheetsButton_Click(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            // get header
            var header = _target.Worksheet.Range["A1"].Resize[4, _target.Columns.Count];

            // find all group label by finding where column B is blank but column C is not
            var labels = ((Range) _target.Columns["B"].Cells).Cast<Range>().Select(cell => cell.Value2)
                .Distinct().Where(value => value != null);

            // create a diction to store the starting address of each label
            var dic = new Dictionary<string, string>();
            // foreach label, create a new sheet and copy to new sheet.
            foreach (var label in labels)
            {
                // add a new worksheet
                var sheet = (_target.Worksheet.Parent as Workbook)?.Sheets.Add() as Worksheet;
                sheet.Name = label;

                // find correspond content
                var content = _target.Rows.Cast<Range>().Where(row =>
                    row.Columns["B"].Value2 == label || row.Columns["C"].Value2 == label);

                // add location info to dictionary
                dic.Add(label, (content.FirstOrDefault().Cells[1, 1] as Range).get_Address());

                // write to the child sheet
                // XXX: not work with linq
                var i = 1;
                foreach (var item in content) item.Copy(sheet.Rows[i++]);
            }

            // update link and add header for children sheets
            foreach (var sheet in ((Workbook) _target.Worksheet.Parent).Sheets.Cast<Worksheet>()
                .Where(sheet => labels.Contains(sheet.Name)))
            {
                sheet.UsedRange.Copy();
                _target.Worksheet.Activate();
                // paste format
                _target.Worksheet.Range[dic[sheet.Name]].PasteSpecial(XlPasteType.xlPasteFormats,
                    XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                // paste content
                // XXX: pastespecial with link not work
                _target.Worksheet.Paste(Link: true);

                // add header for child sheet
                header.Copy();
                //sheet.Activate();
                ((Range) sheet.Rows[1]).Insert(XlInsertShiftDirection.xlShiftDown);

                // adjust row height and column width
                sheet.UsedRange.Rows.AutoFit();
                sheet.UsedRange.Columns.AutoFit();
            }
        }

        #endregion

        #region Private Properties

        private Range _target;

        // Declares at class level to avoid being garbage collected
        private Office.CommandBarButton _generateSheetsButton;
        private Office.CommandBarButton _generateChildrenSheetsButton;

        private dynamic ss;

        #endregion
    }
}