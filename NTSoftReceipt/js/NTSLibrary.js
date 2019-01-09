String.prototype.replaceAll = function (strTarget, strSubString) {
    var strText = this;
    if (strText.length > 0) {
        var intIndexOfMatch = strText.indexOf(strTarget);
        while (intIndexOfMatch != -1) {
            strText = strText.replace(strTarget, strSubString)
            intIndexOfMatch = strText.indexOf(strTarget);
        }
        return (strText);
    }
    else {
        return "";
    }
}
function HienThiControl(item, visibe) {
    var container = document.getElementById(item);
    if (visibe) {
        container.style.visibility = "visible";
        container.style.display = "Block";
    }
    else {
        container.style.visibility = "hidden";
        container.style.display = "none";
    }
}
function checkDate(value,flag)
{
    var re = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
    //Cho phép rỗng
    if (flag == true)
    {
        if (!value.match(re)) {
            return false;
        }
        else
            return true;
    }
    else if (value != '' && !value.match(re)) { 
        return false;
    }
}
//Kiểm tra value null hoặc bằng 0
function checkEmtyValue(value)
{
    if (value == "0" || value == "")
        return true;
    return false;
}
//Kiểm tra value null
function isEmty(value) {
    if (value == "")
        return true;
    return false;
}
//Định dạng window
function dinhDangWindow(winDowName) {
    var screenWidth = screen.width;
    var screenHeight = screen.height;
    var WindowSize = winDowName.getSize();
    winDowName.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(WindowSize.width)) / 2), window.scrollY + 100);
    winDowName.Open();
}
//Định dạng số
function dinhDangSo(sender, value) {
    sender.value(formatNumber(value));
}
//Định dạng tiền tệ
function formatNumber(str) {
    debugger;
    if (jQuery.type(str) == "undefined" || str==null)
        return '0';
    str = str.toString();
    var m = str.lastIndexOf(",");
    var phanNguyen = "";
    var phanTapPhan = "";
    if (m == -1)
        phanNguyen = str;
    else {
        phanNguyen = str.substring(0, str.lastIndexOf(","));
        phanTapPhan = str.substring(str.lastIndexOf(","), str.length);
    }
    var kq = "";
    var dem = 0;
    phanNguyen = parseFloat(phanNguyen.split(".").join('').split(",").join('')).toString();
    for (var i = phanNguyen.length; i > 0; i--) {

        if (!isNaN(phanNguyen.substring(i, i - 1))) {
            kq = phanNguyen.substring(i, i - 1).toString() + kq;
            if (dem == 2 && i != 1) {
                kq = "." + kq;
                dem = 0;
            } else {
                dem = dem + 1;
            }
        }
    }
    if (phanTapPhan != '' && phanTapPhan != ',' && parseFloat(phanTapPhan.split(",").join('').split(".").join('')).toString() != "0") {
        phanTapPhan = parseFloat(phanTapPhan.split(",").join('').split(".").join('')).toString();
        phanTapPhan = ',' + phanTapPhan;
    }
    else {
        phanTapPhan = '';
    }
    if (kq + phanTapPhan == '')
        return 0;

    kq = kq + phanTapPhan;
    return kq;
}
 
function performSearch(grid, index, value) {
    for (var i = index; i < grid.ColumnsCollection.length - 1; i++) {
        var s = grid.ColumnsCollection[i].DataField;
        if (grid.ColumnsCollection[i].Visible == true) {
            grid.addFilterCriteria(s, OboutGridFilterCriteria.Contains, value);
        }
    }
    grid.executeFilter();
    searchTimeout = null;
    return false;
}
var searchTimeout = null;
function searchValue(grid, index, value) {
    debugger;
    if (searchTimeout != null) {
        return false;
    } 
    if (jQuery.type(value) == "undefined")
        value = '';
    for (var i = index; i < grid.ColumnsCollection.length; i++) {
        var s = grid.ColumnsCollection[i].DataField; 
        if (grid.ColumnsCollection[i].Visible == true && s!="") {
            grid.addFilterCriteria(s, OboutGridFilterCriteria.Contains, value);
        }
    }
    //if (jQuery.type(grid.executeFilter()) == "undefined") {
    //    alert("Looix");
    //    return false;
    //}
    searchTimeout = window.setTimeout(grid.executeFilter(), 2000);
    searchTimeout = null;
    return false;
}

function bieuDoTron(result,tenBieuDo,DivID)
{  
    var dataTableGoogle = new google.visualization.DataTable(); 
    for (var i = 0; i < result.Columns.length; i++) {
        if (i == 0)
            dataTableGoogle.addColumn('string', result.Columns[i].Name);
        else
            dataTableGoogle.addColumn('number', result.Columns[i].Name);
    }
    dataTableGoogle.addRows(
                JSON.parse(result.toJSON()).Rows
            );
    // Instantiate and draw our chart, passing in some options
    var chart = new google.visualization.PieChart(document.getElementById(DivID));
    chart.draw(dataTableGoogle,
        {
            title: tenBieuDo,
            position: "top",
            fontsize: "14px",
            chartArea: { width: '100%', height: '100%' },
        });
}
function bieuDoTron3D(result, tenBieuDo, DivID) { 
    var dataTableGoogle = new google.visualization.DataTable();
    for (var i = 0; i < result.Columns.length; i++) {
        if (i == 0)
            dataTableGoogle.addColumn('string', result.Columns[i].Name);
        else
            dataTableGoogle.addColumn('number', result.Columns[i].Name);
    }
    dataTableGoogle.addRows(
                JSON.parse(result.toJSON()).Rows
            );
    // Instantiate and draw our chart, passing in some options
    //var chart = new google.visualization.PieChart(document.getElementById(DivID));
    var options = {
        title: tenBieuDo,
                position: "top",
                fontsize: "12px",
                top:'20px',
                chartArea: { width: '100%', height: '80%' },
                is3D:true
    };

    var chart = new google.visualization.PieChart(document.getElementById(DivID));

    chart.draw(dataTableGoogle, options);

    //chart.draw(dataTableGoogle,
    //    {
    //        title: tenBieuDo,
    //        position: "top",
    //        fontsize: "14px",
    //        top:'20px',
    //        chartArea: { width: '300px', height: '100%' },
    //        is3D:true
    //    });
}
function veBieuDoCot(result, tenBieuDo, DivID) {
     
    var dataTableGoogle = new google.visualization.DataTable();
    for (var i = 0; i < result.Columns.length; i++) {
        if (i == 0)
            dataTableGoogle.addColumn('string', result.Columns[i].Name);
        else
            dataTableGoogle.addColumn('number', result.Columns[i].Name);
    }
    dataTableGoogle.addRows(
                JSON.parse(result.toJSON()).Rows
            );
    // Instantiate and draw our chart, passing in some options
    var chart = new google.visualization.ColumnChart(document.getElementById(DivID));
    chart.draw(dataTableGoogle,
        {
            title: tenBieuDo,
            position: "top",
            fontsize: "14px", 
        });
}
function veBieuDoPhatTrien(result, tenBieuDo, DivID) { 
    if (result == null) {
        $('#' + DivID).html("");

    }
    else {
        var dataTableGoogle = new google.visualization.DataTable();
        for (var i = 0; i < result.Columns.length; i++) {
            if (i == 0)
                dataTableGoogle.addColumn('string', result.Columns[i].Name);
            else
                dataTableGoogle.addColumn('number', result.Columns[i].Name);
        }
        dataTableGoogle.addRows(
                    JSON.parse(result.toJSON()).Rows
                );
        var options = {
            title: tenBieuDo, 
        };
        var chart = new google.visualization.LineChart(document.getElementById(DivID));
        chart.draw(dataTableGoogle, options);
    }
}