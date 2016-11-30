#region Namespaces

using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.UI;
using System.Web.UI.WebControls;
using Sitecore.Data;
using Sitecore.Data.Fields;
using Sitecore.Data.Items;
using Sitecore.SecurityModel; 

#endregion

namespace ExcelTransferUtility.sitecore_modules.Web.Excel_Transfer_Utility.Common
{
    public class UtilityMethods
    {
        #region Fields

        private readonly Database _master = Database.GetDatabase("master");

        #endregion
        
        #region Utility Methods

        /// <summary>
        ///     Gets the mapped drop-down from the user selection
        /// </summary>
        /// <param name="controlCollection"></param>
        /// <returns></returns>
        public DropDownList GetMappedDropdown(IEnumerable controlCollection)
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

        /// <summary>
        ///     Creates the Sitecore Item in the Content Tree
        /// </summary>
        /// <param name="parentItemId"></param>
        /// <param name="itemName"></param>
        /// <param name="templateId"></param>
        /// <returns></returns>
        public Item CreateSitecoreItem(string parentItemId, string itemName, string templateId)
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

        /// <summary>
        ///     Sets simple text field
        /// </summary>
        /// <param name="newItem"></param>
        /// <param name="sitecoreFieldName"></param>
        /// <param name="sitecoreValue"></param>
        public void SetSimpleField(BaseItem newItem, string sitecoreFieldName, string sitecoreValue)
        {
            newItem.Fields[sitecoreFieldName].Value = sitecoreValue;
        }

        /// <summary>
        ///     Sets Link Field
        /// </summary>
        /// <param name="newItem"></param>
        /// <param name="sitecoreFieldName"></param>
        /// <param name="sitecoreValue"></param>
        public void SetLinkField(BaseItem newItem, string sitecoreFieldName, string sitecoreValue)
        {
            LinkField linkField = newItem.Fields[sitecoreFieldName];

            if (linkField == null)
                return;

            linkField.Url = sitecoreValue;
        }

        /// <summary>
        ///     Sets Checkbox field
        /// </summary>
        /// <param name="newItem"></param>
        /// <param name="sitecoreFieldName"></param>
        public void SetCheckboxField(BaseItem newItem, string sitecoreFieldName)
        {
            CheckboxField checkboxField = newItem.Fields[sitecoreFieldName];

            if (checkboxField == null)
                return;

            checkboxField.Checked = true;
        }

        /// <summary>
        ///     Sets MultilistField
        /// </summary>
        /// <param name="item"></param>
        /// <param name="sitecoreFieldName"></param>
        /// <param name="sitecoreValue"></param>
        public void AddToMultiListField(BaseItem item, string sitecoreFieldName, string sitecoreValue)
        {
            MultilistField multilistField = item.Fields[sitecoreFieldName];
            multilistField?.Add(sitecoreValue);
        }

        /// <summary>
        ///     Gets the MultilistField values
        /// </summary>
        /// <param name="list"></param>
        /// <returns></returns>
        public string GetMultiListFieldValues(IEnumerable<Item> list)
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

        /// <summary>
        ///     Gets Link Field values
        /// </summary>
        /// <param name="linkField"></param>
        /// <returns></returns>
        public string GetLinkFieldValue(LinkField linkField)
        {
            var str = string.Empty;

            if (!string.IsNullOrWhiteSpace(linkField?.Url))
                str = linkField.Url;

            return str;
        }

        /// <summary>
        ///     Gets the Checkbox Field value
        /// </summary>
        /// <param name="checkBoxFieldValue"></param>
        /// <returns></returns>
        public string GetCheckBoxFieldValue(bool? checkBoxFieldValue)
        {
            var str = "false";

            if (checkBoxFieldValue.HasValue && checkBoxFieldValue.GetValueOrDefault(false))
                str = "true";

            return str;
        }

        #endregion
    }
}