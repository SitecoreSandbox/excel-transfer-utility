<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExcelTransferUtility.aspx.cs" Inherits="ExcelTransferUtility.sitecore_modules.Web.Excel_Transfer_Utility.ExcelTransferUtility" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Excel Transfer Utility</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <!-- Latest compiled and minified CSS -->
    <link rel="stylesheet" href="//netdna.bootstrapcdn.com/bootstrap/3.0.3/css/bootstrap.min.css" />
    <!-- Optional theme -->
    <link rel="stylesheet" href="//netdna.bootstrapcdn.com/bootstrap/3.0.3/css/bootstrap-theme.min.css" />
    <!-- Custom styles for this template -->
    <link href="Content/css/ExcelTransferUtility.css" rel="stylesheet" />
</head>
<body>
    <form id="Form1" runat="server">

    <div class="container">
        <div class="utility">
            <h1 style="text-align: center;">Excel Transfer Utility</h1>

            <%-- Transfer Selection  --%>
            <asp:Panel ID="PnlTransferSelection" runat="server" CssClass="utility-panel">
                <h4>Choose desired action:</h4>
                <p>
                    <asp:Button ID="BtnImportSelection" runat="server" Text="Import" OnClick="ImportSelection_Click" />
                </p>
                <p>
                    <asp:Button ID="BtnExportSelection" runat="server" Text="Export" OnClick="ExportSelection_Click" />
                </p>
            </asp:Panel>

            <%-- Import Utility --%>
            <asp:Panel ID="PnlImportUtility" runat="server" Visible="false" CssClass="utility-panel">

                <%--Upload--%>
                <asp:Panel ID="PnlUpload" runat="server">
                    <h4>Requirements:</h4>
                    <ul>
                        <li>Microsoft Excel 97-2003 files only (Newer Excel files must be converted to 97-2003 before uploading)</li>
                        <li>Headers are in the first row of the spreadsheet</li>
                        <li>Item names are in the first column of the spreadsheet</li>
                    </ul>

                    <h4>Please select an Excel file from your computer to import:</h4>
                    <p>
                        <asp:FileUpload ID="FuFileToImport" runat="server" CssClass="fileUpload" />
                    </p>
                    <p>
                        <asp:Button runat="server" ID="BtnUpload" Text="Upload" OnClick="Upload_Click" />
                    </p>
                </asp:Panel>

                <%--Import Selection--%>
                <asp:Panel ID="PnlImportSelection" runat="server" Visible="false">
                    <h4>Select sheet name to import:</h4>
                    <p>
                        <asp:DropDownList ID="DdlSheetNames" runat="server"></asp:DropDownList>
                    </p>
                    <br />
                    <h4>Enter the Template ID to use for the imported items:</h4>
                    <p>Example: {D2775315-00DC-4CF4-8B68-E9748127D188}</p>
                    <p>
                        <asp:TextBox ID="TxtTemplateId" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RfvTemplateId" runat="server" Text="*" ErrorMessage="Enter template ID" ControlToValidate="TxtTemplateId"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="RevTemplateId" runat="server" Text="*" ErrorMessage="Enter valid template ID" ControlToValidate="TxtTemplateId" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="BtnNextImportSelection" Text="Next" OnClick="NextImportSelection_Click" />
                    </p>                    
                </asp:Panel>

                <%--Import Mapping & Import--%>
                <asp:Panel ID="PnlImportMapping" runat="server" Visible="false">
                    <asp:ListView ID="LvImportMapping" runat="server" OnItemDataBound="ImportMapping_ItemDataBound">
                        <LayoutTemplate>
                            <h4>Select the fields where excel data will be imported to:</h4>
                            <table>
                                <tr>
                                    <td>
                                        <strong>Spreadsheet Fields</strong>
                                    </td>
                                    <td>
                                        <strong>Template Item Fields</strong>
                                    </td>
                                </tr>
                                <asp:PlaceHolder ID="itemPlaceholder" runat="server" />
                            </table>
                        </LayoutTemplate>
                        <ItemTemplate>
                            <tr>
                                <td>
                                    <asp:Label ID="lblImportFrom" runat="server"></asp:Label>
                                </td>
                                <td>
                                    <asp:DropDownList ID="ddlImportTo" runat="server"></asp:DropDownList>
                                </td>
                            </tr>
                        </ItemTemplate>
                    </asp:ListView>
                    <h4>Enter the Parent Item ID (Import items will be imported as children of this item):</h4>
                    <p>Example: {A91FD891-C477-45F5-B20A-7CFA7F8B53E5}</p>
                    <p>
                        <asp:TextBox ID="TxtParentItemId" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RfvParentId" runat="server" Text="*" ErrorMessage="Enter parent item ID" ControlToValidate="TxtParentItemId"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="RevParentId" runat="server" Text="*" ErrorMessage="Enter valid parent item ID" ControlToValidate="TxtParentItemId" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="BtnImport" Text="Import" OnClick="Import_Click" />
                    </p>
                    <asp:HiddenField ID="HdnFileName" runat="server" />
                    <asp:HiddenField ID="HdnTemplateId" runat="server" />
                    <asp:HiddenField ID="HdnSheetName" runat="server" />
                </asp:Panel>
            </asp:Panel>

            <%-- Export Utility --%>
            <asp:Panel ID="PnlExportUtility" runat="server" Visible="false" CssClass="utility-panel">

                <%--Export Selection--%>
                <asp:Panel runat="server" ID="PnlExportSelection">
                    <h4>Select an option for export:</h4>
                    <asp:RadioButtonList ID="RblExportSelection" runat="server" RepeatLayout="UnorderedList">
                        <asp:ListItem Text="Single Item" Value="single"></asp:ListItem>
                        <asp:ListItem Text="Multiple Items" Value="multiple"></asp:ListItem>
                    </asp:RadioButtonList>
                    <p>
                        <asp:Button runat="server" ID="BtnExportSelectionNext" Text="Next" OnClick="ExportSelectionNext_Click" />
                    </p>
                </asp:Panel>

                <%--Export Single Item--%>
                <asp:Panel runat="server" ID="PnlExportSingle" Visible="false">
                    <h4>Enter the Item ID to export to .csv file:</h4>
                    <p>Example: {923FF496-F233-4152-A6A8-D9964B08C54E}</p>
                    <p>
                        <asp:TextBox ID="TxtExportSingle" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RfvExportSingle" runat="server" Text="*" ErrorMessage="Enter item ID" ControlToValidate="TxtExportSingle"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="RevExportSingle" runat="server" Text="*" ErrorMessage="Enter valid item ID" ControlToValidate="TxtExportSingle" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="BtnExportSingle" Text="Export" OnClick="ExportSingle_Click" />
                    </p>
                </asp:Panel>

                <%--Export Multiple Items--%>
                <asp:Panel runat="server" ID="PnlExportMultiple" Visible="false">
                    <h4>Enter Parent Item ID (Children items will be exported to .csv file):</h4>
                    <p>Example: {A91FD891-C477-45F5-B20A-7CFA7F8B53E5}</p>
                    <p>
                        <asp:TextBox ID="TxtExportMultiple" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RfvExportMultiple" runat="server" Text="*" ErrorMessage="Enter parent item ID" ControlToValidate="TxtExportMultiple"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="RevExportMultiple" runat="server" Text="*" ErrorMessage="Enter valid parent item ID" ControlToValidate="TxtExportMultiple" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="BtnExportMultiple" Text="Export" OnClick="ExportMultiple_Click" />
                    </p>
                </asp:Panel>

            </asp:Panel>

            <%-- Status --%>
            <asp:Panel runat="server" ID="PnlStatus" CssClass="utility-panel">
                <p>
                    <asp:Label runat="server" ID="LblStatus" />
                </p>
                <p>
                    <asp:ValidationSummary ID="VsSummaries" CssClass="validation-summary" DisplayMode="BulletList" EnableClientScript="true" runat="server"/>
                </p>
            </asp:Panel>

        </div>
    </div>

    </form>
    <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script type="text/javascript" src="https://code.jquery.com/jquery.js"></script>
    <!-- Latest compiled and minified JavaScript -->
    <script type="text/javascript" src="//netdna.bootstrapcdn.com/bootstrap/3.0.3/js/bootstrap.min.js"></script>
</body>
</html>
