'use strict';

ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");
//sampleeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee
function initializePage() {
    var context = SP.ClientContext.get_current();
    var user = context.get_web().get_currentUser();

    // This code runs when the DOM is ready and creates a context object which is needed to use the SharePoint object model
    $(document).ready(function () {
        // getUserName();
    });

    // This function prepares, loads, and then executes a SharePoint query to get the current users information
    function getUserName() {
        context.load(user);
        context.executeQueryAsync(onGetUserNameSuccess, onGetUserNameFail);
    }
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;

    function onGetUserNameSuccess() {
        $('#message').text('Hello ' + user.get_title());
        $('#fortag').text(getDataByListInternalName('Months'));

        getDataByListInternalName("months").done(function (item) {
            if (item.d.results.length > 0) {
                for (var i = 0; i <= item.d.results.length; i++) {
                    $('<option value="' + item.d.results[i].Title + '">' + item.d.results[i].Title + '</option>')
                        .appendTo('#Monthdroplist');
                }
            }

        }).fail(function (err) {
            //showError(spService.parseAjaxError(err));
        });
    }

    // This function is executed if the above call fails
    function onGetUserNameFail(sender, args) {
        alert('Failed to get user name. Error:' + args.get_message());
    }


    function getDataByListInternalName(listName) {
        debugger;
        console.log("In Get Name");
        var endpoint = "/_api/web/lists/getbytitle('" + listName + "')/Items?";
        endpoint = siteUrl + endpoint;
        endpoint += "&$expand=AttachmentFiles";
        var ajax = $.ajax({
            url: endpoint,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        return ajax;
    }






}
function getMonthlist(listName) {
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;
    var endpoint = "/_api/web/lists/getbytitle('" + listName + "')/Items?";
    endpoint = siteUrl + endpoint;
    endpoint += "&$expand=AttachmentFiles";
    return endpoint;
}



var SpForms = function (formId) {


    var $form = $(formId);
    const monthList = "months";
    const sectionList = "sections";
    const sectionDetailsList = "sectiondetails";
    var siteUrl = _spPageContextInfo.webAbsoluteUrl;
    var letterwidth = "1000px";
    function getDataByListInternalName(listName, id = 0) {
        var endpoint = "/_api/web/lists/getbytitle('" + listName + "')/Items?$select=";
        endpoint = siteUrl + endpoint;
        if (id !== 0) {
            endpoint += "month / Title, section / Title,month/Volume,*";
            endpoint += "&$filter=(month/Id eq '" + id + "')";
            endpoint += "&$expand=month,section";
        }
        else {
            endpoint += "*";
        }
        var ajax = $.ajax({
            url: endpoint,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" }
        });
        return ajax;
    }

    var data = {};

    function getFullName(listName) {
        var endpoint = "/_api/web/lists/getbytitle('" + listName + "')/Items?select=*,months/Title,sections/Title";
        endpoint += "&$expand=months,sections";
        endpoint = siteUrl + endpoint;
        var ajax = $.ajax({
            url: endpoint,
            method: "GET",
            headers: { "Accept": "application/json;odata=verbose" }
        });
        return ajax;
    }

    function createTable(data) {
        let mastertable = document.getElementById('lettertable');
        let mastertr = document.createElement('TR');
        //mastertd.colspan = data.length;
        data.forEach(function (item) {
            let mastertd = document.createElement('TD');
            if (data.length === 1) {
                mastertd.colSpan = 2;
            }
            let innertable = document.createElement('TABLE');
            innertable.width = "100%";
            innertable.cellSpacing = 0;
            innertable.cellPadding = 0;
            let innertr = document.createElement('TR');
            if (item.image != null) {
                let td = document.createElement('TD');
                td.rowSpan = 4;
                let image = document.createElement('img');
                image.setAttribute("src", item.image.Url);
                let imageprop = JSON.parse(item.imageprop);
                image.height = imageprop.height;
                image.width = imageprop.width;
                innertable.setAttribute("style", "background-color:" + imageprop.colour);
                td.append(image);
                innertr.append(td);
                let sectiontd = document.createElement("TD");
                sectiontd.height = "50";
                sectiontd.align = "center";
                let sectionPTag = document.createElement("P");
                sectionPTag.align = "left";
                sectionPTag.setAttribute("style", "border-radius:25px;padding: 8.5pt 8.5pt 8.5pt 8.5pt");
                let spanForSpace = document.createElement("SPAN");
                for (let i = 0; i < 4; i++)  spanForSpace.innerHTML += "&nbsp";
                sectionPTag.append(spanForSpace);
                let spanforTitle = document.createElement("SPAN");
                spanforTitle.setAttribute("style", "background-color: #343351; border-radius:25px; color: white; font-family: 'Lucida Sans', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif; font-size: 10.0pt;  padding: 3.8pt 20pt  3.8pt");
                spanforTitle.innerText = item.section.Title;
                sectionPTag.append(spanforTitle);
                sectiontd.append(sectionPTag);
                innertr.append(sectiontd);
                innertable.append(innertr);

            }


            if (item.textheading != null) {
                let tr = document.createElement('TR');
                let td = document.createElement('TD');
                td.setAttribute("style", "padding:8.5px 8.5px 10.5px 8.5px");
                let h4 = document.createElement('H4');
                h4.setAttribute("style", item.textheadingprop);
                h4.textContent = item.textheading;
                td.append(h4);
                tr.append(td);
                innertable.append(tr);
            }

            if (item.text) {
                let tr = document.createElement('TR');
                let td = document.createElement('TD');
                td.setAttribute("style", "padding:8.5px 8.5px 10.5px 8.5px");
                let p = document.createElement('P');
                p.setAttribute("style", "margin-left:10px;margin-right:10px");
                p.textContent = item.text;
                p.setAttribute("style", item.textprop);
                td.append(p);
                tr.append(td);
                innertable.append(tr);

                //for read more tag
                let readMoreTr = document.createElement("TR");
                let readMoreTd = document.createElement("TD");
                readMoreTd.colSpan = 2;
                let readMorePTag = document.createElement("P");
                readMorePTag.align = "right";
                readMorePTag.textContent = "READ MORE..";
                readMorePTag.setAttribute("style", "color:white;margin-right:30px;margin-bottom:5px;font-family: 'Lucida Sans', 'Lucida Sans Regular', 'Lucida Grande', 'Lucida Sans Unicode', Geneva, Verdana, sans-serif;");
                readMoreTd.append(readMorePTag);
                readMoreTr.append(readMoreTd);
                innertable.append(readMoreTr);
            }
            mastertd.append(innertable);
            mastertr.append(mastertd);
        });
        mastertable.append(mastertr);
    }

    function groupBy(items, propertyName) {
        var result = [];
        $.each(items, function (index, item) {
            if ($.inArray(item[propertyName], result) == -1) {
                result.push(item[propertyName]);
                //result.push(item);
            }
        });
        return result;
    }

    function run() {
        $.when(
            //get data from entity list
            getDataByListInternalName(monthList).done(function (item) {
                if (item.d.results.length > 0) {
                    for (var i = 0; i < item.d.results.length; i++) {
                        $('<option value="' + item.d.results[i].Id + '">' + item.d.results[i].Title + '</option>')
                            .appendTo('#Monthdroplist');
                    }
                }

            })

        ).then(function () {
            console.log("To check");
        });

        var btn = $form.find('.btn-submit');
        var cpy = $form.find('.btn-copy');
        var generatestatus = false;
        cpy.on('click', function () {
            if (generatestatus) {
                navigator.clipboard.writeText(document.getElementById("letter").innerHTML);
                alert("Code Copied");
            }
            else {
                alert("Generate HTML First to copy");
            }
        });
        btn.on('click', function () {
            generatestatus = true;
            let selectedMonthId = document.getElementById("Monthdroplist").value;
            $("#lettertable").html("");
            let table = document.getElementById("lettertable");
            let titlerow = document.createElement("TR");
            titlerow.setAttribute("style", "text-align:center;background-color:white");
            let td = document.createElement("TD");
            td.colSpan = 2;
            let mwlogo = document.createElement("img");
            mwlogo.setAttribute("src", "https://motifworksinc.sharepoint.com/:i:/r/sites/Home/SiteAssets/SitePages/202001-Newsletter-Vol-9/motiflogo2.png?csf=1&e=M9MG64");
            mwlogo.height = 100;
            mwlogo.width = 300;
            td.append(mwlogo);
            titlerow.append(td);
            table.append(titlerow);

            getFullName(sectionDetailsList).done(function (item) {
                debugger;
                console.log(item);
            });


            getDataByListInternalName(sectionDetailsList, selectedMonthId).done(function (item) {
                data = item.d.results;

                //For Border image
                let imgrow = document.createElement("TR");
                imgrow.setAttribute("style", "text-align:center;background-color:white");
                let imgtd = document.createElement("TD");
                imgtd.colSpan = 2;
                imgtd.width = letterwidth;
                let img = document.createElement("img");
                img.setAttribute("src", "https://motifworksinc.sharepoint.com/:i:/r/sites/Home/SiteAssets/SitePages/202001-Newsletter-Vol-9/77206-image.png?csf=1&e=ByMjYz");
                img.height = 30;
                img.width = 914;
                imgtd.append(img);
                imgrow.append(imgtd);
                table.append(imgrow);

                //for Monthtitle
                let monthTR = document.createElement("TR");
                monthTR.setAttribute("style", "background-color:white");
                let monthtitleTD = document.createElement("TD");
                monthtitleTD.width = 500;
                monthtitleTD.align = "left";
                let pTag = document.createElement("P");
                pTag.align = "center";
                pTag.setAttribute("style", "text-align:center");
                let span = document.createElement("span");
                span.setAttribute("style", "font-size:36.0pt;color:#404040;font-family:Calibri");
                span.innerHTML = "Monthly Newsletter";
                pTag.append(span);
                monthtitleTD.append(pTag);
                monthTR.append(monthtitleTD);

                //for sideimg
                let monthImgTD = document.createElement("TD");
                monthImgTD.align = "right";
                monthImgTD.setAttribute("style", "padding: 8.5pt 8.5pt 8.5pt 8.5pt");
                let monthimg = document.createElement("img");
                monthimg.setAttribute("src", "https://motifworksinc.sharepoint.com/:i:/r/sites/Home/SiteAssets/SitePages/202001-Newsletter-Vol-9/71010-image.png?csf=1&e=f5sDXj");
                monthimg.height = 150;
                monthimg.width = 440;
                monthImgTD.append(monthimg);
                monthTR.append(monthImgTD);

                table.append(monthTR);

                //for volume row
                let volumeTr = document.createElement("TR");
                let volumeTd = document.createElement("TD");
                volumeTd.colSpan = 2;
                let italic = document.createElement("I");
                volumeTd.setAttribute("style", "background:#F2F2F2");
                volumeTd.height = 30;
                let spanForSpace = document.createElement("span");
                for (let i = 0; i < 50; i++)  spanForSpace.innerHTML += "&nbsp";
                volumeTd.append(spanForSpace);
                let volSpan = document.createElement("span");
                volSpan.setAttribute("style", "color:#404040;font-family:11.0pt;font-family:Calibri;");
                volSpan.textContent = data[0].month.Volume + ", " + data[0].month.Title;
                italic.append(volSpan);
                volumeTd.append(italic);
                volumeTr.append(volumeTd);
                table.append(volumeTr);

                //grouping by row
                let group = data.reduce((r, a) => {
                    r[a.rownumber] = [...r[a.rownumber] || [], a];
                    return r;
                }, {});
                var categoryNames = groupBy(data, 'rownumber');
                console.log("group", group);
                console.log("categoryNames", categoryNames);
                categoryNames.forEach(function (item) {
                    createTable(group[item]);
                });
                //footer
                let foorterTr = document.createElement("TR");
                foorterTr.setAttribute("style", "height = 195.6pt");
                foorterTr.innerHTML = ;
                table.append(foorterTr);
            }).fail(function (err) { })


        });
    }

    return {
        run: run
    }
}

