using System;
using System.Linq;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Office = Microsoft.Office.Core;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

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

            // add a new summary sheet
            Worksheet ws = (_target.Worksheet.Parent as Workbook)?.Sheets.Add();
            ws.Name = "汇总";

            // find all group label by finding where column B is blank but column C is not
            var labels = ((Range) _target.Columns["B"].Cells).Cast<Range>().Select(cell => cell.Value2)
                .Distinct().Where(value => value != null);

            // foreach label, create a new sheet and copy to new sheet.
            foreach (var label in labels)
            {
                // add a new worksheet
                var sheet = (_target.Worksheet.Parent as Workbook)?.Sheets.Add() as Worksheet;
                sheet.Name = label;

                // add content to that sheet
                var content = _target.Rows.Cast<Range>().Where(row =>
                    row.Columns["B"].Value2 == label || row.Columns["C"].Value2 == label);
                // XXX: not work with linq
                var i = 1;
                foreach (var item in content) item.Copy(sheet.Rows[i++]);

                // paste to a new summary sheet
                sheet.UsedRange.Copy();
                ws.Activate();
                // paste format
                ws.Range[(content.FirstOrDefault().Cells[1, 1] as Range).get_Address()].PasteSpecial(XlPasteType.xlPasteFormats,
                    XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                // paste content
                // XXX: pastespecial with link not work
                ws.Paste(Link: true);

                // add header for child sheet
                header.Copy();
                //sheet.Activate();
                ((Range)sheet.Rows[1]).Insert(XlInsertShiftDirection.xlShiftDown);
            }

            // Add header
            ws.Activate();
            ws.Range["A1"].Select();
            ws.PasteSpecial();
        }

        #endregion

        #region Private Properties

        private Range _target;

        // Declares at class level to avoid being garbage collected
        private Office.CommandBarButton _generateChildrenSheetsButton;

        #endregion
    }
}