using System;
using System.Collections.Generic;
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
            _generateChildrenSheetsButton = (Office.CommandBarButton) Globals.AddInTechnology.Application
                .CommandBars["cell"].Controls.Add(
                    Office.MsoControlType.msoControlButton,
                    missing, missing, missing, true);
            _generateChildrenSheetsButton.BeginGroup = true;
            _generateChildrenSheetsButton.Caption = "生成子表";
            _generateChildrenSheetsButton.Click += _generateChildrenSheetsButton_Click;
            _generateChildrenSheetsButton.Move(Before: 1);
        }

        private void AddInTechnology_Shutdown(object sender, EventArgs e)
        {
            // remove the control
            Application.CommandBars["Cell"].Reset();

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

                // adjust column width
                sheet.UsedRange.Columns.AutoFit();
            }
        }

        #endregion

        #region Private Properties

        private Range _target;

        // Declares at class level to avoid being garbage collected
        private Office.CommandBarButton _generateChildrenSheetsButton;

        #endregion
    }
}