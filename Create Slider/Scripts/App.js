'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

function initializePage()
{
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        $('#siteContentLink').attr('href', decodeURIComponent(getQueryStringParameter('SPHostUrl')) + "/_layouts/15/viewlsts.aspx");
        $('#createList').click(createSliderList);
    });

    //check for Slider List
    function checkSliderList() {
        $.ajax({
            url: decodeURIComponent(getQueryStringParameter('SPAppWebUrl') + '/_api/SP.AppContextSite(@target)/web/lists/?$Filter=Title eq \'Slider\'&@target=\'' + getQueryStringParameter('SPHostUrl') + '\''),
            method: "GET",
            headers: {
                "Accept": "application/json; odata=verbose"
            },
            success: function (data) {
                if (data.d.results.length == 0) {
                    //continue
                    createSliderList()
                } else {
                    console.info('There is already a SharePoint list with the name Slider is installed on the hostWeb.');
                    alert('There is already a SharePoint list with a name of Slider. Cannot create two lists with the same name.');
                }
            },
            error: function (err) {
                console.error(err);
            }
        });
    }

    //#1 Create List
    function createSliderList() {
        var imgTarget = decodeURIComponent(getQueryStringParameter('SPHostUrl')) + '/SiteAssets/i_Slider.png'

        // Create an announcement SharePoint list with the name that the user specifies.
        var hostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        var currentcontext = new SP.ClientContext.get_current();
        var hostcontext = new SP.AppContextSite(currentcontext, hostUrl);
        var hostweb = hostcontext.get_web();

        //Set ListCreationInfomation()
        var listCreationInfo = new SP.ListCreationInformation();
        listCreationInfo.set_title('Slider');
        listCreationInfo.set_templateType(SP.ListTemplateType.genericList);
        var newList = hostweb.get_lists().add(listCreationInfo);
        newList.set_imageUrl(imgTarget);
        newList.update();

        //Set column data
        var newCols = [
            "<Field Type='Note' DisplayName='Body' Required='FALSE' EnforceUniqueValues='FALSE' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' StaticName='Body' Name='Body'/>",
            "<Field Type='URL' DisplayName='Background Image' Required='FALSE' EnforceUniqueValues='FALSE' Format='Hyperlink' StaticName='BackgroundImage' Name='BackgroundImage'/>",
            "<Field Type='Boolean' DisplayName='Enabled' EnforceUniqueValues='FALSE' StaticName='Enabled' Name='Enabled'><Default>1</Default></Field>",
            "<Field Type='DateTime' DisplayName='Expire' Required='FALSE' EnforceUniqueValues='FALSE' Format='DateTime' FriendlyDisplayFormat='Disabled' StaticName='Expire' Name='Expire'/>",
            "<Field Type='Number' DisplayName='Order' Required='FALSE' EnforceUniqueValues='FALSE' StaticName='Order0' Name='Order0'/>"
        ];
        var newListWithColumns;
        for (var i = 0; i < newCols.length; i++) {
            newListWithColumns = newList.get_fields().addFieldAsXml(newCols[i], true, SP.AddFieldOptions.defaultValue);
        }

        //final load/execute
        context.load(newListWithColumns);
        context.executeQueryAsync(function () {
            //Slider list created successfully!
            uploadTileImage();
            createConfigList();
            alert('Slider list created successfully!');
        },
        function (sender, args) {
            console.error(sender);
            console.error(args);
            if (args.get_message() == 'A list, survey, discussion board, or document library with the specified title already exists in this Web site.  Please choose another title.') {
                alert('The Slider is already installed on the site. SharePoint cannot create another Slider instance on the same subsite.')
            } else {
                alert('Failed to create the Slider list. ' + args.get_message());
            }
        });
    }

    //#2a Upload Tile Image
    function uploadTileImage() {
        BinaryUpload.Uploader().Upload("/images/i_Slider.png", "/SiteAssets/i_Slider.png");
        BinaryUpload.Uploader().Upload("/images/l_Slider.png", "/SiteAssets/l_Slider.png");
    }
    
    //#2b create spConfig List
    function createConfigList() {
        var hostUrl = decodeURIComponent(getQueryStringParameter("SPHostUrl"));
        var currentcontext = new SP.ClientContext.get_current();
        var hostcontext = new SP.AppContextSite(currentcontext, hostUrl);
        var hostweb = hostcontext.get_web();

        //Set ListCreationInfomation()
        var listCreationInfo = new SP.ListCreationInformation();
        listCreationInfo.set_title('spConfig');
        listCreationInfo.set_templateType(SP.ListTemplateType.genericList);
        var newList = hostweb.get_lists().add(listCreationInfo);
        newList.set_hidden(true);
        newList.set_onQuickLaunch(false);
        newList.update();

        //Set column data
        var newListWithColumns = newList.get_fields().addFieldAsXml("<Field Type='Note' DisplayName='Value' Required='FALSE' EnforceUniqueValues='FALSE' NumLines='6' RichText='TRUE' RichTextMode='FullHtml' StaticName='Value' Name='Value'/>", true, SP.AddFieldOptions.defaultValue);

        //final load/execute
        context.load(newListWithColumns);
        context.executeQueryAsync(function () {
            //spConfig list created successfully!
            findInSideNav();
        },
        function (sender, args) {
            console.error(sender);
            console.error(args);
            alert('Failed to create the spConfig list. ' + args.get_message());
        });
    }

    //#3b Find spConfig in side nav
    function findInSideNav() {
        $.ajax({
            url: decodeURIComponent(getQueryStringParameter('SPAppWebUrl') + '/_api/SP.AppContextSite(@target)/web/navigation/quicklaunch?$filter=Title eq \'Recent\'&$expand=Children&$select=Title,Children/Title,Children/Id&@target=\'' + getQueryStringParameter('SPHostUrl') + '\''),
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            success: function (data) {
                //only run if it found one
                if (data.d.results.length == 1) {
                    $.each(data.d.results[0].Children.results, function (index, value) {
                        //find spConfig
                        if (value.Title == "spConfig") {
                            removeInSideNav(value.Id);
                            return false;
                        }
                    })
                }
            },
            error: function (err) {
                console.error(err);
            }
        });
    }

    //#4b remove spConfig in sideNav
    function removeInSideNav(id) {
        $.ajax({
            url: decodeURIComponent(getQueryStringParameter('SPAppWebUrl') + '/_api/SP.AppContextSite(@target)/web/navigation/quicklaunch(' + id + ')?@target=\'' + getQueryStringParameter('SPHostUrl') + '\''),
            method: "DELETE",
            headers: {
                "Accept": "application/json; odata=verbose",
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            },
            success: function (data) {
            },
            error: function (err) {
                console.error(err);
            }
        });
    }

}

function getQueryStringParameter(param) {
    var params = document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == param) {
            return singleParam[1];
        }
    }
}
