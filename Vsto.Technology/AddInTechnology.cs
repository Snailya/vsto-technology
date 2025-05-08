using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Vsto.Technology.Helper;
using Vsto.Technology.Properties;
using Office = Microsoft.Office.Core;

namespace Vsto.Technology;

public partial class AddInTechnology
{
    private void AddInTechnology_Startup(object sender, EventArgs e)
    {
        // add a new handler on a sheet before right-click
        Globals.AddInTechnology.Application.SheetBeforeRightClick += ApplicationOnSheetBeforeRightClick;

        _sampleButton = (Office.CommandBarButton)Globals.AddInTechnology.Application
            .CommandBars["cell"].Controls.Add(
                Office.MsoControlType.msoControlButton,
                missing, missing, missing, true);
        _sampleButton.Tag = "sample";
        _sampleButton.Caption = "样例";
        _sampleButton.Click += SampleButtonOnClick;
        _sampleButton.Move(Before: 1);

        // add a new control, if temporary is false, the functionality may lost though the button is still there after close and then reopen
        _generateSheetsButton = (Office.CommandBarButton)Globals.AddInTechnology.Application
            .CommandBars["cell"].Controls.Add(
                Office.MsoControlType.msoControlButton,
                missing, missing, missing, true);
        _generateSheetsButton.Tag = "gsb";
        _generateSheetsButton.Caption = "生成表";
        _generateSheetsButton.Click += _generateSheetsButton_Click;
        _generateSheetsButton.Move(Before: 1);

        _generateChildrenSheetsButton = (Office.CommandBarButton)Globals.AddInTechnology.Application
            .CommandBars["cell"].Controls.Add(
                Office.MsoControlType.msoControlButton,
                missing, missing, missing, true);
        _generateChildrenSheetsButton.Tag = "gcsb";
        _generateChildrenSheetsButton.Caption = "生成子表";
        _generateChildrenSheetsButton.Click += _generateChildrenSheetsButton_Click;
        _generateChildrenSheetsButton.Move(Before: 2);
    }

    private void SampleButtonOnClick(Office.CommandBarButton ctrl, ref bool cancelDefault)
    {
        SampleCreator.Create(_target);
    }

    private string EnsureTemplateExist()
    {
        const string templateName = "XX涂装项目报价表模板.xltx";

        string templatePath;
        
        try
        {
            var defaultLocalFileLocation = Globals.AddInTechnology.Application.DefaultFilePath;

            var customOfficeTemplatePath = Path.Combine(defaultLocalFileLocation, "Custom Office Templates");
            if (!Directory.Exists(customOfficeTemplatePath))
            {
                var customOfficeTemplatePathZh = Path.Combine(defaultLocalFileLocation, "自定义Office模板");
                if (!Directory.Exists(customOfficeTemplatePathZh)) Directory.CreateDirectory(customOfficeTemplatePath);
            }

            // 检查是否存在templates
            templatePath = Path.Combine(customOfficeTemplatePath, templateName);

            if (!File.Exists(templatePath)) File.WriteAllBytes(templatePath, Resources.Template);
        }
        catch (RegistryValueNotFoundException exception)
        {
            // if you find a template path failed, prompt the user to select the template path
            var openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            openFileDialog.Filter = @"Excel template files (*.xltx)|*.xltx";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            openFileDialog.RestoreDirectory = true;

            // Show the dialog and check the result
            if (openFileDialog.ShowDialog() == DialogResult.OK)
                templatePath = openFileDialog.FileName; // For a single file
            else
                throw new TemplateNotFoundException("No template selected by the user.");
        }

        return templatePath;
    }

    private bool TryGetPersonalTemplatePathFromRegistry(out string personalTemplatePath)
    {
        personalTemplatePath = string.Empty;

        try
        {
            personalTemplatePath = GetPersonalTemplatesPath();
            return true;
        }
        catch (Exception e)
        {
            // ignored
        }

        return false;
    }


    private void AddInTechnology_Shutdown(object sender, EventArgs e)
    {
        // remove the control
        Globals.AddInTechnology.Application.CommandBars["cell"].FindControl(Tag: "gsb")?.Delete();
        Globals.AddInTechnology.Application.CommandBars["cell"].FindControl(Tag: "gcsb")?.Delete();

        // remove a sheet before right click handler
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

    private static string GetPersonalTemplatesPath()
    {
        string value;

        try
        {
            // first try to find registry value from the current user
            value = RegistryHelper.GetRegistryValue(Registry.CurrentUser,
                @"Software\Microsoft\Office\16.0\Excel\Options",
                "PersonalTemplates");

            return value;
        }
        catch (RegistryValueNotFoundException)
        {
            try
            {
                value = RegistryHelper.GetRegistryValue(Registry.LocalMachine,
                    @"Software\Microsoft\Office\16.0\Excel\Options",
                    "PersonalTemplates");

                return value;
            }
            catch (RegistryValueNotFoundException e)
            {
                throw new RegistryValueNotFoundException(
                    "Unable to find personal templates path from registry key.");
            }
        }
        catch (Exception e)
        {
            throw new RegistryValueNotFoundException(
                $"Unable to find personal templates path from registry key. {e.Message}");
        }
    }


    #region Callbacks

    private void ApplicationOnSheetBeforeRightClick(object sh, Range target, ref bool cancel)
    {
        // store target
        _target = target;
        
        // todo: 拆分后的子表中还是可以插入样例，原始表中没选择数据的时候也会显示生成表，需要详细控制下。
        if (target.Worksheet.Name == "原始表")
        {
            _sampleButton.Visible = false;
            _generateChildrenSheetsButton.Visible = true;
        }
        else
        {
            _sampleButton.Visible = true;
            _generateChildrenSheetsButton.Visible = false;
        }
    }

    /// <summary>
    ///     Generates quotation sheets by user selection using templates
    /// </summary>
    /// <param name="ctrl"></param>
    /// <param name="canceldefault"></param>
    private void _generateSheetsButton_Click(Office.CommandBarButton ctrl, ref bool canceldefault)
    {
        var templatePath = EnsureTemplateExist();

        // find all group labels by finding where column B is blank but column C is not
        var labels = ((Range)_target.Columns["B"].Cells).Cast<Range>().Select(cell => cell.Value2 as string)
            .Distinct().Where(value => value != null);
        // create a diction to store the starting address of each label
        var sectionRowLookup = ((Range)_target.Columns["B"].Cells).Cast<Range>()
            .ToLookup(c => c.Value2 as string, c => c.Row - _target[1, 1].Row as int?);

        // add a new workbook
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
                        var rng = (Range)sheet.Cells[row, 1].Resize[1, sheet.UsedRange.Columns.Count];
                        rng.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent5;
                        rng.Interior.TintAndShade = 0.5;
                    });
                    continue;
                }

                // write summary
                var sumFormulaSource = (Range)sheet.Cells[absoluteRowGroup.Min() - 1, 1];
                sumFormulaSource.Formula = $"=SUM(A{absoluteRowGroup.Min()}:A{absoluteRowGroup.Max()})";

                var sumDestination = sheet.Range[sumFormulaSource.Address,
                    sheet.Cells[sumFormulaSource.Row, sheet.UsedRange.Columns.Count]];

                sumFormulaSource.AutoFill(sheet.Range[sumFormulaSource.Address, sumDestination.Address],
                    XlAutoFillType.xlFillValues);

                ((Range)sheet.Rows[$"{absoluteRowGroup.Min()}:{absoluteRowGroup.Max()}"]).Group();
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
        var labels = ((Range)_target.Columns["B"].Cells).Cast<Range>().Select(cell => cell.Value2)
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
        foreach (var sheet in ((Workbook)_target.Worksheet.Parent).Sheets.Cast<Worksheet>()
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
            ((Range)sheet.Rows[1]).Insert(XlInsertShiftDirection.xlShiftDown);

            // adjust row height and column width
            sheet.UsedRange.Rows.AutoFit();
            sheet.UsedRange.Columns.AutoFit();
        }
    }

    #endregion

    #region Private Properties

    private Range _target;

    // Declares at class level to avoid being garbage collected
    private Office.CommandBarButton _sampleButton;
    private Office.CommandBarButton _generateSheetsButton;
    private Office.CommandBarButton _generateChildrenSheetsButton;

    #endregion
}