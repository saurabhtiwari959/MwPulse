<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>

<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <SharePoint:ScriptLink Name="sp.js" runat="server" OnDemand="true" LoadAfterUI="true" Localizable="false" />
    <meta name="WebPartPageExpansion" content="full" />

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/App.js"></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    Page Title
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">

    <div>
        <p id="message"></p>
        <a href="JavaScript:window.location = _spPageContextInfo.webAbsoluteUrl + '/Lists/months/AllItems.aspx';">Months</a>
        <a href="JavaScript:window.location = _spPageContextInfo.webAbsoluteUrl + '/Lists/sectiondetails/AllItems.aspx';">Section Details</a>
        <a href="JavaScript:window.location = _spPageContextInfo.webAbsoluteUrl + '/Lists/sections/AllItems.aspx';">Sections</a>
        <p id="fortag"></p>

        <div data-list-name="ListItem" id="ListItemForm">
            <div>
                <select id="Monthdroplist" name="Months">Select months</select>

                <input type="button" class="btn-submit" value="Generate" />
                <input type="button" class="btn-copy" value="Copy" />
            </div>
            <div id="letter">
                <table cellpadding="0" cellspacing="0" align="center" style="background-color: #D0CECE; width: 1000px;">
                    <tr>
                        <td id="lettertd" style="padding: 8.5pt 8.5pt 0pt 8.5pt">
                            <table id="lettertable" cellpadding="0" cellspacing="0" align="center" style="background-color: white;">
                            </table>
                        </td>
                    </tr>

                </table>

            </div>

        </div>

        <script type="text/javascript">
            $(document).ready(function () {
                var spForm = SpForms('#ListItemForm');
                spForm.run();
            });
        </script>

    </div>


</asp:Content>
