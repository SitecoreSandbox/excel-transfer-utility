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
        protected DropDownList DdlItemName;
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

        protected void ImportSelection_Click(object sender, EventArgs e)
        {
            PnlImportUtility.Visible = true;
            PnlTransferSelection.Visible = false;
        }

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
                DdlItemName.DataSource = dictionary.Values.ToList();
                DdlItemName.DataBind();
            }
            catch (Exception ex)
            {
                LblStatus.Text = $"Error: {ex.Message}";
            }
        }

        protected void ImportMapping_ItemDataBound(object sender, ListViewItemEventArgs e)
        {
            if (e.Item.ItemType != ListViewItemType.DataItem)
                return;
            var listViewDataItem = e.Item as ListViewDataItem;
            if (listViewDataItem == null)
                return;
            var str = listViewDataItem.DataItem as string;
            var dropDownList = listViewDataItem.FindControl("ddlImportTo") as DropDownList;
            var label = listViewDataItem.FindControl("lblImportFrom") as Label;
            if (dropDownList == null || label == null || string.IsNullOrWhiteSpace(str))
                return;
            label.Text = str;
            dropDownList.Attributes.Add("importId", str);
            var path = HdnTemplateId.Value;
            if (string.IsNullOrWhiteSpace(path))
                return;
            var obj1 = _master.GetItem(path);
            if (obj1 == null)
                return;
            var list =
                obj1.Axes.GetDescendants()
                    .Where(descendant => descendant.TemplateID.ToString() == "{455A3E98-A627-4B40-8035-E683A0331AC7}")
                    .ToList();
            if (!list.Any())
                return;
            foreach (var obj2 in list)
                dropDownList.Items.Add(new ListItem(obj2.Name, obj2.Name));
        }

        protected void Import_Click(object sender, EventArgs e)
        {
            try
            {
                var str1 = TxtTemplateId.Text.Trim();
                var fileName = HdnFileName.Value;
                if (string.IsNullOrWhiteSpace(fileName) || string.IsNullOrWhiteSpace(str1))
                    return;
                var file = new FileInfo(fileName);
                var workbook = Workbook.getWorkbook(file);
                var sheet = workbook.getSheet(HdnSheetName.Value);
                var parentItemId = TxtParentItemId.Text.Trim();
                var dictionary = new Dictionary<int, string>();

                for (var row = 0; row < sheet.getRows(); ++row)
                {
                    var selectedValue = DdlItemName.SelectedValue;
                    Item newItem = null;

                    for (var index = 0; index < sheet.getColumns(); ++index)
                    {
                        if (row == 0)
                        {
                            dictionary.Add(index, sheet.getCell(index, row).getContents());
                        }
                        else
                        {
                            var str2 = dictionary[index].ToLower();
                            if (newItem == null)
                                newItem = CreateSitecoreItem(parentItemId, selectedValue, str1);
                            var list1 = LvImportMapping.Items.ToList();
                            if (!list1.Any()) continue;

                            foreach (var control in list1)
                            {
                                var mappedDropdown = GetMappedDropdown(control.Controls);
                                if (mappedDropdown == null || mappedDropdown.Attributes["importId"].ToLower() != str2)
                                    continue;
                                var sitecoreFieldName = mappedDropdown.SelectedValue;
                                var sitecoreFieldType = string.Empty;
                                var contents = sheet.getCell(index, row).getContents();
                                var obj1 = _master.GetItem(str1);

                                if (obj1 != null)
                                {
                                    var list2 =
                                        obj1.Axes.GetDescendants()
                                            .Where(
                                                descendant =>
                                                    descendant.TemplateID.ToString() ==
                                                    Constants.FieldIds.TemplateFieldId)
                                            .ToList();
                                    if (list2.Any())
                                    {
                                        var obj2 =
                                            list2.FirstOrDefault(item => item.Name == sitecoreFieldName);
                                        if (obj2 != null)
                                            sitecoreFieldType = obj2["Type"];
                                    }
                                }

                                UpdateSitecoreFields(newItem, sitecoreFieldName, sitecoreFieldType, contents);
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

        protected void ExportSelection_Click(object sender, EventArgs e)
        {
            PnlTransferSelection.Visible = false;
            PnlExportUtility.Visible = true;
        }

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

        private void ExportToCsv(IReadOnlyCollection<Item> exportItems)
        {
            var csvExport = new CsvExport();

            foreach (var obj1 in exportItems)
            {
                csvExport.AddRow();
                var obj2 = _master.GetItem(obj1.TemplateID);
                if (obj2 == null) continue;
                var list =
                    obj2.Axes.GetDescendants()
                        .Where(
                            descendant =>
                                descendant.TemplateID.ToString() == Constants.FieldIds.TemplateFieldId)
                        .ToList();
                if (!list.Any()) continue;

                foreach (var obj3 in list)
                {
                    var name = obj3.Name;
                    var str = obj1[name];

                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(str)) continue;

                    switch (str)
                    {
                        case "0":
                            str = "false";
                            break;
                        case "1":
                            str = "true";
                            break;
                    }

                    csvExport[name] = str;
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

        private static DropDownList GetMappedDropdown(IEnumerable controlCollection)
        {
            DropDownList dropDownList = null;

            foreach (
                var control in
                    controlCollection.Cast<Control>()
                        .Where(control => control.GetType().ToString() == "System.Web.UI.WebControls.DropDownList"))
            {
                dropDownList = control as DropDownList;
            }

            return dropDownList;
        }

        private Item CreateSitecoreItem(string parentItemId, string itemName, string templateId)
        {
            Item obj1 = null;

            using (new SecurityDisabler())
            {
                var obj2 = _master.GetItem(parentItemId);
                if (obj2 == null) return null;
                TemplateItem template = _master.GetItem(templateId);
                obj1 = obj2.Add(itemName, template);
            }

            return obj1;
        }

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
                            SetSimpleField(newItem, sitecoreFieldName, sitecoreValue);
                        else if (sitecoreFieldType == Constants.FieldIds.CheckboxFieldId && sitecoreValue.ToLower() == "true")
                            SetCheckboxField(newItem, sitecoreFieldName);
                        else if (sitecoreFieldType == Constants.FieldIds.DropListFieldId || sitecoreFieldType == Constants.FieldIds.MultiListFieldId ||
                                 sitecoreFieldType == Constants.FieldIds.TreeListFieldId || sitecoreFieldType == Constants.FieldIds.TreeListExFieldId)
                            AddToMultiListField(newItem, sitecoreFieldName, sitecoreValue);
                        else if (sitecoreFieldType == Constants.FieldIds.GeneralLinkFieldId)
                            SetLinkField(newItem, sitecoreFieldName, sitecoreValue);
                        else
                            SetSimpleField(newItem, sitecoreFieldName, sitecoreValue);
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

        private static void SetSimpleField(BaseItem newItem, string sitecoreFieldName, string sitecoreValue)
        {
            newItem.Fields[sitecoreFieldName].Value = sitecoreValue;
        }

        private static void SetLinkField(BaseItem newItem, string sitecoreFieldName, string sitecoreValue)
        {
            LinkField linkField = newItem.Fields[sitecoreFieldName];

            if (linkField == null)
                return;

            linkField.Url = sitecoreValue;
        }

        private static void SetCheckboxField(BaseItem newItem, string sitecoreFieldName)
        {
            CheckboxField checkboxField = newItem.Fields[sitecoreFieldName];

            if (checkboxField == null)
                return;

            checkboxField.Checked = true;
        }

        private static void AddToMultiListField(BaseItem item, string sitecoreFieldName, string sitecoreValue)
        {
            MultilistField multilistField = item.Fields[sitecoreFieldName];
            multilistField?.Add(sitecoreValue);
        }

        private string GetMultiListFieldValues(IEnumerable<Item> list)
        {
            var stringBuilder = new StringBuilder();
            var enumerable = list as IList<Item> ?? list.ToList();
            if (!enumerable.Any()) return stringBuilder.ToString();
            var num = 1;

            foreach (var obj in enumerable)
            {
                stringBuilder.Append(obj.Name);
                if (num < enumerable.Count())
                    stringBuilder.Append("|");
                ++num;
            }

            return stringBuilder.ToString();
        }

        private string GetLinkFieldValue(LinkField linkField)
        {
            var str = string.Empty;

            if (!string.IsNullOrWhiteSpace(linkField?.Url))
                str = linkField.Url;

            return str;
        }

        private string GetCheckBoxFieldValue(bool? checkBoxFieldValue)
        {
            var str = "false";

            if (checkBoxFieldValue.HasValue && checkBoxFieldValue.GetValueOrDefault(false))
                str = "true";

            return str;
        } 

        #endregion
    }
}