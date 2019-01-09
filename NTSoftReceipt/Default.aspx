<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="NTSoftReceipt._Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
    <style type="text/css">
        .bangmenu{

        }
        .bangmenu td{
            padding:5px;
            font-size:20px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table class="bangmenu">
            <tr>
                <td><a href="danhmuc/dmtinh.aspx">DM tỉnh</a></td>
                <td><a href="danhmuc/dmquanhuyen.aspx">DM quận/huyện</a></td>
                <td><a href="danhmuc/dmxaphuong.aspx">DM xã/phường</a></td>
                <td><a href="danhmuc/dmhuyentinh.aspx">DM Huyện / Tỉnh</a></td>
            </tr>
            <tr>
                <td><a href="danhmuc/dmnhanvien.aspx">DM nhân viên</a></td>
                <td><a href="danhmuc/dmnhapkho.aspx">DM nhập kho</a></td>
                <td><a href="danhmuc/dmnhapkhoNC.aspx">DM nhập kho NC</a></td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
