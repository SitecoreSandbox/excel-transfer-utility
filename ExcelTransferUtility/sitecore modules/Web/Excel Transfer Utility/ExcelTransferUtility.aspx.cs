#region Namespaces

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using CSharpJExcel.Jxl;
using ExcelTransferUtility.sitecore_modules.Web.Excel_Transfer_Utility.Common;
using ExcelTransferUtility.sitecore_modules.Web.Excel_Transfer_Utility.Export;
using Sitecore.Data;
using Sitecore.Data.Fields;
using Sitecore.Data.Items;
using Sitecore.Diagnostics;
using Sitecore.SecurityModel;

#endregion

namespace ExcelTransferUtility.sitecore_modules.Web.Excel_Transfer_Utility
{
    public class ExcelTransferUtility : Page
    {
        #region Fields

        private readonly Database _master = Database.GetDatabase("master");

        #endregion

        #region Page Controls

        protected Button BtnExportMultiple;
        protected Button BtnExportSelection;
        protected Button BtnExportSelectionNext;
        protected Button BtnExportSingle;
        protected Button BtnImport;
        protected Button BtnImportSelection;
        protected Button BtnNextImportSelection;
        protected Button BtnUpload;
        protected DropDownList DdlSheetNames;
        protected HtmlForm Form1;
        protected FileUpload FuFileToImport;
        protected HiddenField HdnFileName;
        protected HiddenField HdnSheetName;
        protected HiddenField HdnTemplateId;
        protected HtmlHead Head1;
        protected Label LblStatus;
        protected ListView LvImportMapping;
        protected Panel PnlExportMultiple;
        protected Panel PnlExportSelection;
        protected Panel PnlExportSingle;
        protected Panel PnlExportUtility;
        protected Panel PnlImportMapping;
        protected Panel PnlImportSelection;
        protected Panel PnlImportUtility;
        protected Panel PnlStatus;
        protected Panel PnlTransferSelection;
        protected Panel PnlUpload;
        protected RadioButtonList RblExportSelection;
        protected RegularExpressionValidator RevExportMultiple;
        protected RegularExpressionValidator RevExportSingle;
        protected RegularExpressionValidator RevParentId;
        protected RegularExpressionValidator RevTemplateId;
        protected RequiredFieldValidator RfvExportMultiple;
        protected RequiredFieldValidator RfvExportSingle;
        protected RequiredFieldValidator RfvParentId;
        protected RequiredFieldValidator RfvTemplateId;
        protected TextBox TxtExportMultiple;
        protected TextBox TxtExportSingle;
        protected TextBox TxtParentItemId;
        protected TextBox TxtTemplateId;
        protected ValidationSummary VsSummaries;

        #endregion

        #region Page Control Methods

        /// <summary>
        /// When the Import button is clicked upon
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ImportSelection_Click(object sender, EventArgs e)
        {
            PnlImportUtility.Visible = true;
            PnlTransferSelection.Visible = false;
        }

        /// <summary>
        /// Uploads Excel file to be used for Import
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Upload_Click(object sender, EventArgs e)
        {
            if (!FuFileToImport.HasFile)
                return;

            try
            {
                if ((FuFileToImport.PostedFile.ContentType == "application/vnd.ms-excel" ||
                     FuFileToImport.PostedFile.ContentType ==
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") &&
                    (FuFileToImport.PostedFile.FileName.Contains(".xlsx") ||
                     FuFileToImport.PostedFile.FileName.Contains(".xls")))
                {
                    var str = Server.MapPath("~/temp/") + Path.GetFileName(FuFileToImport.FileName);
                    FuFileToImport.SaveAs(str);
                    HdnFileName.Value = str;
                    DdlSheetNames.DataSource = Workbook.getWorkbook(new FileInfo(str)).getSheetNames();
                    DdlSheetNames.DataBind();
                    LblStatus.Text = string.Empty;
                    PnlUpload.Visible = false;
                    PnlImportSelection.Visible = true;
                }
                else
                    LblStatus.Text = "Upload Status: Only Microsoft 97-2003 Excel files are accepted";
            }
            catch (Exception ex)
            {
                LblStatus.Text = "Upload status: The file could not be uploaded. The following error occured: " +
                                 ex.Message;
            }
        }

        /// <summary>
        /// Allows you select sheet name and template ID for imported items
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void NextImportSelection_Click(object sender, EventArgs e)
        {
            try
            {
                HdnTemplateId.Value = TxtTemplateId.Text.Trim();
                var fileName = HdnFileName.Value;
                if (string.IsNullOrEmpty(fileName))
                    return;
                var workbook = Workbook.getWorkbook(new FileInfo(fileName));
                var sheet = workbook.getSheet(DdlSheetNames.SelectedValue);
                HdnSheetName.Value = DdlSheetNames.SelectedValue;
                var dictionary = new Dictionary<int, string>();
                for (var row = 0; row < sheet.getRows(); ++row)
                {
                    for (var index = 0; index < sheet.getColumns(); ++index)
                    {
                        if (row == 0)
                            dictionary.Add(index, sheet.getCell(index, row).getContents());
                    }
                }
                workbook.close();
                LblStatus.Text = string.Empty;
                PnlImportSelection.Visible = false;
                PnlImportMapping.Visible = true;
                LvImportMapping.DataSource = dictionary.Values.ToList();
                LvImportMapping.DataBind();
            }
            catch (Exception ex)
            {
                LblStatus.Text = $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Gets the fields associated with the template specified
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ImportMapping_ItemDataBound(object sender, ListViewItemEventArgs e)
        {
            if (e.Item.ItemType == ListViewItemType.DataItem)
            {
                var listViewDataItem = e.Item as ListViewDataItem;

                if (listViewDataItem != null && listViewDataItem.DataItemIndex > 0)
                {
                    var str = listViewDataItem.DataItem as string;
                    var dropDownList = listViewDataItem.FindControl("ddlImportTo") as DropDownList;
                    var label = listViewDataItem.FindControl("lblImportFrom") as Label;

                    if (dropDownList != null && label != null && !string.IsNullOrWhiteSpace(str))
                    {
                        label.Text = str;
                        dropDownList.Attributes.Add("importId", str);

                        var path = HdnTemplateId.Value;

                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            var obj1 = _master.GetItem(path);

                            if (obj1 != null)
                            {
                                var list =
                                    obj1.Axes.GetDescendants()
                                        .Where(
                                            descendant =>
                                                descendant.TemplateID.ToString() ==
                                                "{455A3E98-A627-4B40-8035-E683A0331AC7}")
                                        .ToList();

                                if (list.Any())
                                {
                                    foreach (var obj2 in list)
                                    {
                                        dropDownList.Items.Add(new ListItem(obj2.Name, obj2.Name));
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    var str = listViewDataItem.DataItem as string;
                    var dropDownList = listViewDataItem.FindControl("ddlImportTo") as DropDownList;
                    var label = listViewDataItem.FindControl("lblImportFrom") as Label;

                    if (dropDownList != null && label != null && !string.IsNullOrWhiteSpace(str))
                    {
                        label.Text = str;
                        dropDownList.Visible = false;
                    }
                }
            }
        }

        /// <summary>
        /// Performs the Import into Sitecore
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void Import_Click(object sender, EventArgs e)
        {
            try
            {
                // Trim the template ID
                var templateId = TxtTemplateId.Text.Trim();

                // Get the file name value
                var fileName = HdnFileName.Value;

                // Ensure strings are present before proceeding
                if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(templateId))
                    return;

                // Prepare for Import
                var file = new FileInfo(fileName);
                var workbook = Workbook.getWorkbook(file);
                var sheet = workbook.getSheet(HdnSheetName.Value);
                var parentItemId = TxtParentItemId.Text.Trim();
                var dictionary = new Dictionary<int, string>();

                // Perform the Import
                // Foreach row...
                for (var row = 0; row < sheet.getRows(); ++row)
                {
                    Item newItem = null;
                    
                    // Foreach column in this row...
                    for (var index = 0; index < sheet.getColumns(); ++index)
                    {
                        // If this is the header row...
                        if (row == 0)
                        {
                            dictionary.Add(index, sheet.getCell(index, row).getContents());
                        }
                        // Else if it is not the header row..
                        else
                        {
                            if (newItem == null)
                            {
                                newItem = new UtilityMethods().CreateSitecoreItem(parentItemId, sheet.getCell(0, row).getContents(), templateId);
                            }

                            var lvImportMappingList = LvImportMapping.Items.ToList();

                            if (lvImportMappingList.Any())
                            {
                                // Don't get the control for the first column in the spreadsheet as that is the item name
                                var i = 0;

                                foreach (var control in lvImportMappingList)
                                {
                                    if (i > 0)
                                    {
                                        var mappedDropdown = new UtilityMethods().GetMappedDropdown(control.Controls);
                                        var importId = dictionary[index].ToLower();

                                        if (mappedDropdown != null &&
                                            mappedDropdown.Attributes["importId"].ToLower() == importId)
                                        {
                                            var sitecoreFieldName = mappedDropdown.SelectedValue;
                                            var sitecoreFieldType = string.Empty;
                                            var contents = sheet.getCell(index, row).getContents();
                                            var templateItem = _master.GetItem(templateId);

                                            if (templateItem != null)
                                            {
                                                var fieldList =
                                                    templateItem.Axes.GetDescendants()
                                                        .Where(
                                                            descendant =>
                                                                descendant.TemplateID.ToString() ==
                                                                Constants.FieldIds.TemplateFieldId)
                                                        .ToList();

                                                if (fieldList.Any())
                                                {
                                                    var sitecoreField =
                                                        fieldList.FirstOrDefault(item => item.Name == sitecoreFieldName);
                                                    if (sitecoreField != null)
                                                        sitecoreFieldType = sitecoreField["Type"];
                                                }
                                            }

                                            UpdateSitecoreFields(newItem, sitecoreFieldName, sitecoreFieldType, contents);
                                        }
                                    }

                                    i++;
                                }
                            }
                        }
                    }
                }
                workbook.close();
                file.Delete();
                PnlImportUtility.Visible = false;
                LblStatus.Text = "Excel spreadsheet was imported successfully!";
            }
            catch (Exception ex)
            {
                Log.Error("Error while importing items", ex, (object)ToString());
                LblStatus.Text = $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// When the Import button is clicked upon
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ExportSelection_Click(object sender, EventArgs e)
        {
            PnlTransferSelection.Visible = false;
            PnlExportUtility.Visible = true;
        }

        /// <summary>
        /// When Export button is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ExportSelectionNext_Click(object sender, EventArgs e)
        {
            var selectedValue = RblExportSelection.SelectedValue;

            if (!string.IsNullOrWhiteSpace(selectedValue))
            {
                LblStatus.Text = string.Empty;

                if (selectedValue == "single")
                {
                    PnlExportSelection.Visible = false;
                    PnlExportSingle.Visible = true;
                }
                else
                {
                    if (selectedValue != "multiple")
                        return;
                    PnlExportSelection.Visible = false;
                    PnlExportMultiple.Visible = true;
                }
            }
            else
                LblStatus.Text = "Please make a selection then click the Next button";
        }

        /// <summary>
        /// When Single Item is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ExportSingle_Click(object sender, EventArgs e)
        {
            var exportItems = new List<Item>();
            var str = TxtExportSingle.Text.Trim();

            if (!string.IsNullOrWhiteSpace(str))
            {
                var obj = _master.GetItem(Sitecore.Data.ID.Parse(str));

                if (obj != null)
                {
                    exportItems.Add(obj);

                    if (exportItems.Any())
                    {
                        ExportToCsv(exportItems);
                        PnlExportSingle.Visible = false;
                    }
                    else
                        LblStatus.Text = "No Items are available to export";
                }
                else
                    LblStatus.Text = "Cannot find the item in the content tree";
            }
            else
                LblStatus.Text = "Please enter item ID";
        }

        /// <summary>
        /// When Multiple Items is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void ExportMultiple_Click(object sender, EventArgs e)
        {
            var list = new List<Item>();
            var str = TxtExportMultiple.Text.Trim();

            if (!string.IsNullOrWhiteSpace(str))
            {
                var obj = _master.GetItem(Sitecore.Data.ID.Parse(str));

                if (obj != null)
                {
                    var exportItems = obj.GetChildren().ToList();

                    if (exportItems.Any())
                    {
                        ExportToCsv(exportItems);
                        PnlExportMultiple.Visible = false;
                    }
                    else
                        LblStatus.Text = "No Items are available to export";
                }
                else
                    LblStatus.Text = "Cannot find the parent item in the content tree";
            }
            else
                LblStatus.Text = "Please enter parent item ID";
        }

        #endregion

        #region Private Utility Methods

        /// <summary>
        /// Performs the Exports of item/s
        /// </summary>
        /// <param name="exportItems"></param>
        private void ExportToCsv(IReadOnlyCollection<Item> exportItems)
        {
            var csvExport = new CsvExport();

            foreach (var item in exportItems)
            {
                csvExport.AddRow();

                // Add the item name as the first column in the row
                csvExport["Item Name"] = item.Name;

                // Get the fields in the item and add the data from the item/s
                var templateItem = _master.GetItem(item.TemplateID);

                if (templateItem != null)
                {
                    var fieldList =
                        templateItem.Axes.GetDescendants()
                            .Where(
                                descendant =>
                                    descendant.TemplateID.ToString() == Constants.FieldIds.TemplateFieldId)
                            .ToList();

                    if (fieldList.Any())
                    {
                        foreach (var field in fieldList)
                        {
                            var fieldName = field.Name;
                            var fieldValue = item[fieldName];

                            if (!string.IsNullOrWhiteSpace(fieldName) && !string.IsNullOrWhiteSpace(fieldValue))
                            {
                                switch (fieldValue)
                                {
                                    case "0":
                                        fieldValue = "false";
                                        break;
                                    case "1":
                                        fieldValue = "true";
                                        break;
                                }

                                csvExport[fieldName] = fieldValue;
                            }
                        }
                    }
                }
            }

            try
            {
                var path = $"{Server.MapPath("~/temp/")}{"Export-" + DateTime.Now.ToFileTime()}.csv";
                csvExport.ExportToFile(path);
                LblStatus.Text = File.Exists(path)
                    ? $"<p>Export successful!</p><p>File Location: {path}</p>"
                    : "<p>Items were not exported successfully!</p><p>No file exists in the /temp folder.</p>";
            }
            catch (Exception ex)
            {
                Log.Error("Could not export " + exportItems.Count + "items.", this);
                LblStatus.Text = $"Error: {ex.Message}";
            }
        }

        /// <summary>
        /// Updates the item fields with imported content
        /// </summary>
        /// <param name="newItem"></param>
        /// <param name="sitecoreFieldName"></param>
        /// <param name="sitecoreFieldType"></param>
        /// <param name="sitecoreValue"></param>
        private void UpdateSitecoreFields(Item newItem, string sitecoreFieldName, string sitecoreFieldType,
            string sitecoreValue)
        {
            using (new SecurityDisabler())
            {
                newItem.Editing.BeginEdit();

                try
                {
                    if (!string.IsNullOrWhiteSpace(sitecoreFieldName) &&
                        newItem.Fields[sitecoreFieldName] != null)
                    {
                        if (sitecoreFieldType == Constants.FieldIds.MultiLineTextFieldId || sitecoreFieldType == Constants.FieldIds.RichTextFieldId ||
                            sitecoreFieldType == Constants.FieldIds.SingleLineFieldId || sitecoreFieldType == Constants.FieldIds.DropLinkFieldId ||
                            sitecoreFieldType == Constants.FieldIds.DropListFieldId)
                            new UtilityMethods().SetSimpleField(newItem, sitecoreFieldName, sitecoreValue);
                        else if (sitecoreFieldType == Constants.FieldIds.CheckboxFieldId && sitecoreValue.ToLower() == "true")
                            new UtilityMethods().SetCheckboxField(newItem, sitecoreFieldName);
                        else if (sitecoreFieldType == Constants.FieldIds.DropListFieldId || sitecoreFieldType == Constants.FieldIds.MultiListFieldId ||
                                 sitecoreFieldType == Constants.FieldIds.TreeListFieldId || sitecoreFieldType == Constants.FieldIds.TreeListExFieldId)
                            new UtilityMethods().AddToMultiListField(newItem, sitecoreFieldName, sitecoreValue);
                        else if (sitecoreFieldType == Constants.FieldIds.GeneralLinkFieldId)
                            new UtilityMethods().SetLinkField(newItem, sitecoreFieldName, sitecoreValue);
                        else
                            new UtilityMethods().SetSimpleField(newItem, sitecoreFieldName, sitecoreValue);
                    }
                    newItem.Editing.EndEdit();
                }
                catch (Exception ex)
                {
                    Log.Error("Could not update item " + newItem.Paths.FullPath + ": " + ex.Message, this);
                    LblStatus.Text = $"Error: {ex.Message}";
                    newItem.Editing.CancelEdit();
                }
            }
        }

        #endregion
    }
}