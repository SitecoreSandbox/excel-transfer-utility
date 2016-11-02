<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExcelTransferUtility.aspx.cs" Inherits="Sitecore.Module.ExcelTransferUtility" %>

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
    <link href="~/sitecore modules/Web/Excel Transfer Utility/ExcelTransferUtility.css" rel="stylesheet" />
</head>
<body>
    <form id="form1" runat="server">

    <div class="container">
        <div class="utility">
            <h1 style="text-align: center;">Excel Transfer Utility</h1>

            <%-- Transfer Selection  --%>
            <asp:Panel ID="pnlTransferSelection" runat="server" CssClass="utility-panel">
                <h4>Choose desired action:</h4>
                <p>
                    <asp:Button ID="btnImportSelection" runat="server" Text="Import" OnClick="ImportSelection_Click" />
                </p>
                <p>
                    <asp:Button ID="btnExportSelection" runat="server" Text="Export" OnClick="ExportSelection_Click" />
                </p>
            </asp:Panel>

            <%-- Import Utility --%>
            <asp:Panel ID="pnlImportUtility" runat="server" Visible="false" CssClass="utility-panel">

                <%--Upload--%>
                <asp:Panel ID="pnlUpload" runat="server">
                    <h4>Please select a file from your computer to import:</h4>
                    <p>
                        <asp:FileUpload ID="fuFileToImport" runat="server" CssClass="fileUpload" />
                    </p>
                    <p>
                        <asp:Button runat="server" ID="btnUpload" Text="Upload" OnClick="Upload_Click" />
                    </p>
                </asp:Panel>

                <%--Import Selection--%>
                <asp:Panel ID="pnlImportSelection" runat="server" Visible="false">
                    <h4>Select sheet name to import:</h4>
                    <p>
                        <asp:DropDownList ID="ddlSheetNames" runat="server"></asp:DropDownList>
                    </p>
                    <br />
                    <h4>Enter template ID to use for imported items:</h4>
                    <p>Example: {D2775315-00DC-4CF4-8B68-E9748127D188}</p>
                    <p>
                        <asp:TextBox ID="txtTemplateId" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvTemplateId" runat="server" Text="*" ErrorMessage="Enter template ID" ControlToValidate="txtTemplateId"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revTemplateId" runat="server" Text="*" ErrorMessage="Enter valid template ID" ControlToValidate="txtTemplateId" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="btnNextImportSelection" Text="Next" OnClick="NextImportSelection_Click" />
                    </p>                    
                </asp:Panel>

                <%--Import Mapping & Import--%>
                <asp:Panel ID="pnlImportMapping" runat="server" Visible="false">
                    <h4>Select column name to use for item name:</h4>
                    <p>
                        <asp:DropDownList ID="ddlItemName" runat="server"></asp:DropDownList>
                    </p>
                    <br />                    
                    <asp:ListView ID="lvImportMapping" runat="server" OnItemDataBound="ImportMapping_ItemDataBound">
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
                    <h4>Enter parent item ID (Import items will be imported as children of this item):</h4>
                    <p>Example: {A91FD891-C477-45F5-B20A-7CFA7F8B53E5}</p>
                    <p>
                        <asp:TextBox ID="txtParentItemId" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvParentId" runat="server" Text="*" ErrorMessage="Enter parent item ID" ControlToValidate="txtParentItemId"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revParentId" runat="server" Text="*" ErrorMessage="Enter valid parent item ID" ControlToValidate="txtParentItemId" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="btnImport" Text="Import" OnClick="Import_Click" />
                    </p>
                    <asp:HiddenField ID="hdnFileName" runat="server" />
                    <asp:HiddenField ID="hdnTemplateId" runat="server" />
                    <asp:HiddenField ID="hdnSheetName" runat="server" />
                </asp:Panel>
            </asp:Panel>

            <%-- Export Utility --%>
            <asp:Panel ID="pnlExportUtility" runat="server" Visible="false" CssClass="utility-panel">

                <%--Export Selection--%>
                <asp:Panel runat="server" ID="pnlExportSelection">
                    <h4>Select an option:</h4>
                    <asp:RadioButtonList ID="rblExportSelection" runat="server" RepeatLayout="UnorderedList">
                        <asp:ListItem Text="Single Item" Value="single"></asp:ListItem>
                        <asp:ListItem Text="Multiple Items" Value="multiple"></asp:ListItem>
                    </asp:RadioButtonList>
                    <p>
                        <asp:Button runat="server" ID="btnExportSelectionNext" Text="Next" OnClick="ExportSelectionNext_Click" />
                    </p>
                </asp:Panel>

                <%--Export Single Item--%>
                <asp:Panel runat="server" ID="pnlExportSingle" Visible="false">
                    <h4>Enter item ID to export to .csv file:</h4>
                    <p>Example: {923FF496-F233-4152-A6A8-D9964B08C54E}</p>
                    <p>
                        <asp:TextBox ID="txtExportSingle" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvExportSingle" runat="server" Text="*" ErrorMessage="Enter item ID" ControlToValidate="txtExportSingle"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revExportSingle" runat="server" Text="*" ErrorMessage="Enter valid item ID" ControlToValidate="txtExportSingle" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="btnExportSingle" Text="Export" OnClick="ExportSingle_Click" />
                    </p>
                </asp:Panel>

                <%--Export Multiple Items--%>
                <asp:Panel runat="server" ID="pnlExportMultiple" Visible="false">
                    <h4>Enter parent item ID (Children items will be exported to .csv file):</h4>
                    <p>Example: {A91FD891-C477-45F5-B20A-7CFA7F8B53E5}</p>
                    <p>
                        <asp:TextBox ID="txtExportMultiple" runat="server" CssClass="textbox"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="rfvExportMultiple" runat="server" Text="*" ErrorMessage="Enter parent item ID" ControlToValidate="txtExportMultiple"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="revExportMultiple" runat="server" Text="*" ErrorMessage="Enter valid parent item ID" ControlToValidate="txtExportMultiple" ValidationExpression="^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$"></asp:RegularExpressionValidator>
                    </p>
                    <p>
                        <asp:Button runat="server" ID="btnExportMultiple" Text="Export" OnClick="ExportMultiple_Click" />
                    </p>
                </asp:Panel>

            </asp:Panel>

            <%-- Status --%>
            <asp:Panel runat="server" ID="pnlStatus" CssClass="utility-panel">
                <p>
                    <asp:Label runat="server" ID="lblStatus" />
                </p>
                <p>
                    <asp:ValidationSummary ID="vsSummaries" CssClass="validation-summary" DisplayMode="BulletList" EnableClientScript="true" runat="server"/>
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
