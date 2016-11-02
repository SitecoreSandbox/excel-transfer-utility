using CSharpJExcel.Jxl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Sitecore.Data;
using Sitecore.Data.Items;
using Sitecore.Data.Fields;
using Sitecore.Diagnostics;
using Sitecore.SecurityModel;

namespace Sitecore.Module
{
    public class ExcelTransferUtility : Page
    {
        protected Button btnExportMultiple;
        protected Button btnExportSelection;
        protected Button btnExportSelectionNext;
        protected Button btnExportSingle;
        protected Button btnImport;
        protected Button btnImportSelection;
        protected Button btnNextImportSelection;
        protected Button btnUpload;
        protected const string checkboxName = "Checkbox";
        protected DropDownList ddlItemName;
        protected DropDownList ddlSheetNames;
        protected const string dropLinkName = "Droplink";
        protected const string dropListName = "Droplist";
        protected const string dropTreeName = "Droptree";
        protected HtmlForm form1;
        protected FileUpload fuFileToImport;
        protected const string generalLinkName = "General Link";
        protected HiddenField hdnFileName;
        protected HiddenField hdnSheetName;
        protected HiddenField hdnTemplateId;
        protected HtmlHead Head1;
        protected Label lblStatus;
        protected ListView lvImportMapping;
        protected Database master = Database.GetDatabase("master");
        protected const string multiLineTextName = "Multi-Line Text";
        protected const string multiListName = "Multilist";
        protected Panel pnlExportMultiple;
        protected Panel pnlExportSelection;
        protected Panel pnlExportSingle;
        protected Panel pnlExportUtility;
        protected Panel pnlImportMapping;
        protected Panel pnlImportSelection;
        protected Panel pnlImportUtility;
        protected Panel pnlStatus;
        protected Panel pnlTransferSelection;
        protected Panel pnlUpload;
        protected RadioButtonList rblExportSelection;
        protected RegularExpressionValidator revExportMultiple;
        protected RegularExpressionValidator revExportSingle;
        protected RegularExpressionValidator revParentId;
        protected RegularExpressionValidator revTemplateId;
        protected RequiredFieldValidator rfvExportMultiple;
        protected RequiredFieldValidator rfvExportSingle;
        protected RequiredFieldValidator rfvParentId;
        protected RequiredFieldValidator rfvTemplateId;
        protected const string richTextName = "Rich Text";
        protected const string singleLineTextName = "Single-Line Text";
        protected const string templateFieldId = "{455A3E98-A627-4B40-8035-E683A0331AC7}";
        protected const string treeListExName = "TreelistEx";
        protected const string treeListName = "Treelist";
        protected TextBox txtExportMultiple;
        protected TextBox txtExportSingle;
        protected TextBox txtParentItemId;
        protected TextBox txtTemplateId;
        protected ValidationSummary vsSummaries;

        private static void AddToMultiListField(Item item, string sitecoreFieldName, string sitecoreValue)
        {
            MultilistField field = item.Fields[sitecoreFieldName];
            if (field != null)
            {
                field.Add(sitecoreValue);
            }
        }

        private Item CreateSitecoreItem(string parentItemId, string itemName, string templateId)
        {
            Item item = null;
            using (new SecurityDisabler())
            {
                Item item2 = this.master.GetItem(parentItemId);
                if (item2 != null)
                {
                    TemplateItem template = this.master.GetItem(templateId);
                    item = item2.Add(itemName, template);
                }
            }
            return item;
        }

        protected void ExportMultiple_Click(object sender, EventArgs e)
        {
            List<Item> source = new List<Item>();
            string str = this.txtExportMultiple.Text.Trim();
            if (!string.IsNullOrWhiteSpace(str))
            {
                Item item = this.master.GetItem(Data.ID.Parse(str));
                if (item != null)
                {
                    source = item.GetChildren().ToList<Item>();
                    if (source.Any<Item>())
                    {
                        this.ExportToCsv(source);
                        this.pnlExportMultiple.Visible = false;
                    }
                    else
                    {
                        this.lblStatus.Text = "No Items are available to export";
                    }
                }
                else
                {
                    this.lblStatus.Text = "Cannot find the parent item in the content tree";
                }
            }
            else
            {
                this.lblStatus.Text = "Please enter parent item ID";
            }
        }

        protected void ExportSelection_Click(object sender, EventArgs e)
        {
            this.pnlTransferSelection.Visible = false;
            this.pnlExportUtility.Visible = true;
        }

        protected void ExportSelectionNext_Click(object sender, EventArgs e)
        {
            string selectedValue = this.rblExportSelection.SelectedValue;
            if (!string.IsNullOrWhiteSpace(selectedValue))
            {
                this.lblStatus.Text = string.Empty;
                if (selectedValue == "single")
                {
                    this.pnlExportSelection.Visible = false;
                    this.pnlExportSingle.Visible = true;
                }
                else if (selectedValue == "multiple")
                {
                    this.pnlExportSelection.Visible = false;
                    this.pnlExportMultiple.Visible = true;
                }
            }
            else
            {
                this.lblStatus.Text = "Please make a selection then click the Next button";
            }
        }

        protected void ExportSingle_Click(object sender, EventArgs e)
        {
            List<Item> source = new List<Item>();
            string str = this.txtExportSingle.Text.Trim();
            if (!string.IsNullOrWhiteSpace(str))
            {
                Item item = this.master.GetItem(Data.ID.Parse(str));
                if (item != null)
                {
                    source.Add(item);
                    if (source.Any<Item>())
                    {
                        this.ExportToCsv(source);
                        this.pnlExportSingle.Visible = false;
                    }
                    else
                    {
                        this.lblStatus.Text = "No Items are available to export";
                    }
                }
                else
                {
                    this.lblStatus.Text = "Cannot find the item in the content tree";
                }
            }
            else
            {
                this.lblStatus.Text = "Please enter item ID";
            }
        }

        private void ExportToCsv(List<Item> exportItems)
        {
            CsvExport export = new CsvExport();
            //foreach (Item item in exportItems)
            //{
            //    export.AddRow();
            //    Item item2 = this.master.GetItem(item.TemplateID);
            //    if (item2 != null)
            //    {
            //        // TODO: Understand, test/debug, fix
            //        if (CS$<> 9__CachedAnonymousMethodDelegate9 == null)
            //        {
            //            CS$<> 9__CachedAnonymousMethodDelegate9 = new Func<Item, bool>(null, (IntPtr) < ExportToCsv > b__8);
            //        }
            //        List<Item> source = item2.Axes.GetDescendants().Where<Item>(CS$<> 9__CachedAnonymousMethodDelegate9).ToList<Item>();
            //        if (source.Any<Item>())
            //        {
            //            foreach (Item item3 in source)
            //            {
            //                string name = item3.Name;
            //                string str2 = item[name];
            //                if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(str2))
            //                {
            //                    if (str2 == "0")
            //                    {
            //                        str2 = "false";
            //                    }
            //                    else if (str2 == "1")
            //                    {
            //                        str2 = "true";
            //                    }
            //                    export[name] = str2;
            //                }
            //            }
            //        }
            //    }
            //}
            try
            {
                string str3 = "Export-" + DateTime.Now.ToFileTime().ToString();
                string path = $"{base.Server.MapPath("~/temp/")}{str3}.csv";
                export.ExportToFile(path);
                if (File.Exists(path))
                {
                    this.lblStatus.Text = $"<p>Export successful!</p><p>File Location: {path}</p>";
                }
                else
                {
                    this.lblStatus.Text = "<p>Items were not exported successfully!</p><p>No file exists in the /temp folder.</p>";
                }
            }
            catch (Exception exception)
            {
                Log.Error("Could not export " + exportItems.Count + "items.", this);
                this.lblStatus.Text = $"Error: {exception.Message}";
            }
        }

        private string GetCheckBoxFieldValue(bool? checkBoxFieldValue)
        {
            string str = "false";
            if (checkBoxFieldValue.HasValue && checkBoxFieldValue.GetValueOrDefault(false))
            {
                str = "true";
            }
            return str;
        }

        private string GetLinkFieldValue(LinkField linkField)
        {
            string url = string.Empty;
            if ((linkField != null) && !string.IsNullOrWhiteSpace(linkField.Url))
            {
                url = linkField.Url;
            }
            return url;
        }

        private DropDownList GetMappedDropdown(ControlCollection controlCollection)
        {
            DropDownList list = null;
            foreach (Control control in controlCollection)
            {
                if (control.GetType().ToString() == "System.Web.UI.WebControls.DropDownList")
                {
                    list = control as DropDownList;
                }
            }
            return list;
        }

        private string GetMultiListFieldValues(IEnumerable<Item> list)
        {
            StringBuilder builder = new StringBuilder();
            if (list.Any<Item>())
            {
                int num = 1;
                foreach (Item item in list)
                {
                    builder.Append(item.Name);
                    if (num < list.Count<Item>())
                    {
                        builder.Append("|");
                    }
                    num++;
                }
            }
            return builder.ToString();
        }

        protected void Import_Click(object sender, EventArgs e)
        {
            try
            {
                string templateId = this.txtTemplateId.Text.Trim();
                string fileName = this.hdnFileName.Value;
                if (!string.IsNullOrWhiteSpace(fileName) && !string.IsNullOrWhiteSpace(templateId))
                {
                    FileInfo file = new FileInfo(fileName);
                    Workbook workbook = Workbook.getWorkbook(file);
                    Sheet sheet = workbook.getSheet(this.hdnSheetName.Value);
                    string parentItemId = this.txtParentItemId.Text.Trim();
                    Dictionary<int, string> dictionary = new Dictionary<int, string>();
                    for (int i = 0; i < sheet.getRows(); i++)
                    {
                        string selectedValue = this.ddlItemName.SelectedValue;
                        Item newItem = null;
                        for (int j = 0; j < sheet.getColumns(); j++)
                        {
                            if (i == 0)
                            {
                                dictionary.Add(j, sheet.getCell(j, i).getContents());
                            }
                            else
                            {
                                string str5 = dictionary[j].ToLower();
                                if (newItem == null)
                                {
                                    newItem = this.CreateSitecoreItem(parentItemId, selectedValue, templateId);
                                }
                                List<ListViewDataItem> source = this.lvImportMapping.Items.ToList<ListViewDataItem>();

                                if (source.Any<ListViewDataItem>())
                                {
                                    foreach (ListViewDataItem item2 in source)
                                    {
                                        DropDownList mappedDropdown = this.GetMappedDropdown(item2.Controls);
                                        if ((mappedDropdown != null) && (mappedDropdown.Attributes["importId"].ToLower() == str5))
                                        {
                                            Func<Item, bool> predicate = null;
                                            string sitecoreFieldName = mappedDropdown.SelectedValue;
                                            string sitecoreFieldType = string.Empty;
                                            string sitecoreValue = sheet.getCell(j, i).getContents();
                                            Item item = this.master.GetItem(templateId);
                                            //if (item != null)
                                            //{
                                            //    // TODO: Understand, test/debug, fix
                                            //    if (item == null)
                                            //    {
                                            //        CS$<> 9__CachedAnonymousMethodDelegate4 = new Func<Item, bool>(null, (IntPtr) < Import_Click > b__2);
                                            //    }

                                            //    List<Item> list3 = item.Axes.GetDescendants().Where<Item>(CS$<> 9__CachedAnonymousMethodDelegate4).ToList<Item>();

                                            //    if (list3.Any<Item>())
                                            //    {
                                            //        if (predicate == null)
                                            //        {
                                            //            <> c__DisplayClass6 class2;
                                            //            predicate = new Func<Item, bool>(class2, (IntPtr)this.< Import_Click > b__3);
                                            //        }

                                            //        Item item4 = list3.Where<Item>(predicate).FirstOrDefault<Item>();

                                            //        if (item4 != null)
                                            //        {
                                            //            sitecoreFieldType = item4["Type"];
                                            //        }
                                            //    }
                                            //}

                                            this.UpdateSitecoreFields(newItem, sitecoreFieldName, sitecoreFieldType, sitecoreValue);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    workbook.close();
                    file.Delete();
                    this.pnlImportUtility.Visible = false;
                    this.lblStatus.Text = "Excel spreadsheet was imported successfully!";
                }
            }
            catch (Exception exception)
            {
                Log.Error("Error while importing items", exception, this.ToString());
                this.lblStatus.Text = $"Error: {exception.Message}";
            }
        }


        protected void ImportMapping_ItemDataBound(object sender, ListViewItemEventArgs e)
        {
            if (e.Item.ItemType == ListViewItemType.DataItem)
            {
                ListViewDataItem item = e.Item as ListViewDataItem;
                if (item != null)
                {
                    string dataItem = item.DataItem as string;
                    DropDownList list = item.FindControl("ddlImportTo") as DropDownList;
                    Label label = item.FindControl("lblImportFrom") as Label;
                    if (((list != null) && (label != null)) && !string.IsNullOrWhiteSpace(dataItem))
                    {
                        label.Text = dataItem;
                        list.Attributes.Add("importId", dataItem);
                        string path = this.hdnTemplateId.Value;
                        if (!string.IsNullOrWhiteSpace(path))
                        {
                            Item item2 = this.master.GetItem(path);
                            if (item2 != null)
                            {
                                //// TODO: Understand, test/debug, fix
                                //if (CS$<> 9__CachedAnonymousMethodDelegate1 == null)
                                //{
                                //    CS$<> 9__CachedAnonymousMethodDelegate1 = new Func<Item, bool>(null, (IntPtr) < ImportMapping_ItemDataBound > b__0);
                                //}

                                //List<Item> source = item2.Axes.GetDescendants().Where<Item>(CS$<> 9__CachedAnonymousMethodDelegate1).ToList<Item>();

                                //if (source.Any<Item>())
                                //{
                                //    foreach (Item item3 in source)
                                //    {
                                //        list.Items.Add(new ListItem(item3.Name, item3.Name));
                                //    }
                                //}
                            }
                        }
                    }
                }
            }
        }

        protected void ImportSelection_Click(object sender, EventArgs e)
        {
            this.pnlImportUtility.Visible = true;
            this.pnlTransferSelection.Visible = false;
        }

        protected void NextImportSelection_Click(object sender, EventArgs e)
        {
            try
            {
                this.hdnTemplateId.Value = this.txtTemplateId.Text.Trim();
                string str = this.hdnFileName.Value;
                if (!string.IsNullOrEmpty(str))
                {
                    FileInfo file = new FileInfo(str);
                    Workbook workbook = Workbook.getWorkbook(file);
                    Sheet sheet = workbook.getSheet(this.ddlSheetNames.SelectedValue);
                    this.hdnSheetName.Value = this.ddlSheetNames.SelectedValue;
                    Dictionary<int, string> dictionary = new Dictionary<int, string>();
                    for (int i = 0; i < sheet.getRows(); i++)
                    {
                        for (int j = 0; j < sheet.getColumns(); j++)
                        {
                            if (i == 0)
                            {
                                dictionary.Add(j, sheet.getCell(j, i).getContents());
                            }
                        }
                    }
                    workbook.close();
                    this.lblStatus.Text = string.Empty;
                    this.pnlImportSelection.Visible = false;
                    this.pnlImportMapping.Visible = true;
                    this.lvImportMapping.DataSource = dictionary.Values.ToList<string>();
                    this.lvImportMapping.DataBind();
                    this.ddlItemName.DataSource = dictionary.Values.ToList<string>();
                    this.ddlItemName.DataBind();
                }
            }
            catch (Exception exception)
            {
                this.lblStatus.Text = $"Error: {exception.Message}";
            }
        }

        private static void SetCheckboxField(Item newItem, string sitecoreFieldName)
        {
            CheckboxField field = newItem.Fields[sitecoreFieldName];
            if (field != null)
            {
                field.Checked = true;
            }
        }

        private static void SetLinkField(Item newItem, string sitecoreFieldName, string sitecoreValue)
        {
            LinkField field = newItem.Fields[sitecoreFieldName];
            if (field != null)
            {
                field.Url = sitecoreValue;
            }
        }

        private static void SetSimpleField(Item newItem, string sitecoreFieldName, string sitecoreValue)
        {
            newItem.Fields[sitecoreFieldName].Value = sitecoreValue;
        }

        private void UpdateSitecoreFields(Item newItem, string sitecoreFieldName, string sitecoreFieldType, string sitecoreValue)
        {
            using (new SecurityDisabler())
            {
                newItem.Editing.BeginEdit();
                try
                {
                    if (((newItem != null) && !string.IsNullOrWhiteSpace(sitecoreFieldName)) && (newItem.Fields[sitecoreFieldName] != null))
                    {
                        if (((sitecoreFieldType == "Multi-Line Text") || (sitecoreFieldType == "Rich Text")) || (((sitecoreFieldType == "Single-Line Text") || (sitecoreFieldType == "Droplink")) || (sitecoreFieldType == "Droptree")))
                        {
                            SetSimpleField(newItem, sitecoreFieldName, sitecoreValue);
                        }
                        else if ((sitecoreFieldType == "Checkbox") && (sitecoreValue.ToLower() == "true"))
                        {
                            SetCheckboxField(newItem, sitecoreFieldName);
                        }
                        else if (((sitecoreFieldType == "Droplist") || (sitecoreFieldType == "Multilist")) || ((sitecoreFieldType == "Treelist") || (sitecoreFieldType == "TreelistEx")))
                        {
                            AddToMultiListField(newItem, sitecoreFieldName, sitecoreValue);
                        }
                        else if (sitecoreFieldType == "General Link")
                        {
                            SetLinkField(newItem, sitecoreFieldName, sitecoreValue);
                        }
                        else
                        {
                            SetSimpleField(newItem, sitecoreFieldName, sitecoreValue);
                        }
                    }
                    newItem.Editing.EndEdit();
                }
                catch (Exception exception)
                {
                    Log.Error("Could not update item " + newItem.Paths.FullPath + ": " + exception.Message, this);
                    this.lblStatus.Text = $"Error: {exception.Message}";
                    newItem.Editing.CancelEdit();
                }
            }
        }

        protected void Upload_Click(object sender, EventArgs e)
        {
            if (this.fuFileToImport.HasFile)
            {
                try
                {
                    if (((this.fuFileToImport.PostedFile.ContentType == "application/vnd.ms-excel") || (this.fuFileToImport.PostedFile.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) && (this.fuFileToImport.PostedFile.FileName.Contains(".xlsx") || this.fuFileToImport.PostedFile.FileName.Contains(".xls")))
                    {
                        string filename = base.Server.MapPath("~/temp/") + Path.GetFileName(this.fuFileToImport.FileName);
                        this.fuFileToImport.SaveAs(filename);
                        this.hdnFileName.Value = filename;
                        FileInfo file = new FileInfo(filename);
                        this.ddlSheetNames.DataSource = Workbook.getWorkbook(file).getSheetNames();
                        this.ddlSheetNames.DataBind();
                        this.lblStatus.Text = string.Empty;
                        this.pnlUpload.Visible = false;
                        this.pnlImportSelection.Visible = true;
                    }
                    else
                    {
                        this.lblStatus.Text = "Upload Status: Only Microsoft 97-2003 Excel files are accepted";
                    }
                }
                catch (Exception exception)
                {
                    this.lblStatus.Text = "Upload status: The file could not be uploaded. The following error occured: " + exception.Message;
                }
            }
        }
    }
}
