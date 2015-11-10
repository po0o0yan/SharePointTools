<%-- The following 4 lines are ASP.NET directives needed when using SharePoint components --%>
PM
<%@ Page Inherits="Microsoft.SharePoint.WebPartPages.WebPartPage, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" MasterPageFile="~masterurl/default.master" Language="C#" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<%-- The markup and script in the following Content element will be placed in the <head> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.min.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.runtime.js"></script>
    <script type="text/javascript" src="/_layouts/15/sp.js"></script>
    <script src="/_layouts/15/sp.requestexecutor.js" type="text/javascript"></script> 

    <!-- Add your CSS styles to the following file -->
    <link rel="Stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Add your JavaScript to the following file -->
    <script type="text/javascript" src="../Scripts/AppHelper.js"></script>
    <script type="text/javascript" src="../Scripts/uploadFile.js"></script>
    <script type="text/javascript"  src="../Scripts/CTypesAndFieldProvisioning.js"></script>
    <script type="text/javascript"  src="../Scripts/manageVariations.js" ></script>
    <script type="text/javascript"  src="../Scripts/CreateSite.js" ></script>
</asp:Content>

<%-- The markup in the following Content element will be placed in the TitleArea of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
     Provisioning
</asp:Content>

<%-- The markup and script in the following Content element will be placed in the <body> of the page --%>
<asp:Content ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <input type="button" value="Publish MasterPage" />
    <input type="button" value="Create List" />
    <input type="button" value="Create Variation Label" />
    <input type="button" value="Create Subsite" />

    <br/>
    <h1>Results</h1>
    <div>
        <p id="message">
        </p>
    </div>
    
    <script type="text/javascript">
        $(":button:First").click(function () {
            //window.PM.uploadFileToHostWeb.execute('/MasterPages/DynacarePortalMasterPage.txt', '_catalogs/masterpage', 'DynacarePortal.master', 'text');
            window.PM.uploadFileToHostWeb.execute('/Images/Capture.png', 'PublishingImages', 'Capture.png', 'binary');
        });

        $(":button:nth-of-type(2)").click(function () {
            window.PM.ProvisionContentTypes.execute();
        });
        $(":button:nth-of-type(3)").click(function () {
            window.PM.createVariationLabel.execute('1033', 'English', 'http://sp/sites/variationroot/en-us', 'en-US', 'English (United States)');
        });
        $(":button:nth-of-type(4)").click(function () {
            window.PM.createSiteInHostWeb.execute("sample","desc","sample","STS#0");
        });
    </script>
</asp:Content>