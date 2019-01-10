<%@ Page Language="C#" MasterPageFile="~/Default.Master" AutoEventWireup="true" CodeBehind="dmquanhuyen.aspx.cs"
    Inherits="NTSoftReceipt.danhmuc.dmquanhuyen" %>

<%@ Register Assembly="obout_ComboBox" Namespace="Obout.ComboBox" TagPrefix="cc3" %>
<%@ Register Assembly="obout_Interface" Namespace="Obout.Interface" TagPrefix="cc2" %>
<%@ Register Assembly="obout_Window_NET" Namespace="OboutInc.Window" TagPrefix="owd" %>
<%@ Register Assembly="obout_Grid_NET" Namespace="Obout.Grid" TagPrefix="cc1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">

    <script type="text/javascript">
    var coThaotac;
    var ma;
    window.onload = function() {
        Grid1.refresh();  
    }
    function  isEmptyText(value)
    {
        var len = value.replace(new  RegExp(" ", "g"), '').length;
        if (len == 0){
            return  true;
        }
        else
        return false;
    }
    function _mLamMoi()
    {
        _txtMa.enable();
        _txtMa.value('');
        _txtTen.value('');
        _txtGhiChu.value('');
        _chkNTD.checked(false);
        _txtTenVT.value('');
        _cboTinh.selectedIndex(0);
    }
    function setWindowPosition() {
            var screenWidth = screen.width;
            var screenHeight = screen.height;
            var Window1Size = Window1.getSize(); 
            Window1.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(Window1Size.width)) / 2), 200);
            Window1.Open();
        }
    function OnBeforeClientAdd()
    {//gan phan quyền
            var permiss = NTSoftReceipt.danhmuc.dmquanhuyen.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác thêm mới!");
                return false;
            }
        coThaotac=1;
        _mLamMoi();
        Window1.setTitle("Thêm Quận/Huyện");
        setWindowPosition();
        return false;
    }
    function OnBeforeClientEdit(record)
    {//gan phan quyền
            var permiss = NTSoftReceipt.danhmuc.dmquanhuyen.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác sửa!");
                return false;
            }
        coThaotac=2;
        _txtMa.disable();
        _txtMa.value(record.maHuyenpr);
        _txtTen.value(record.tenHuyen);
        _txtGhiChu.value(record.ghiChu);
        _txtTenVT.value(record.tenVietTat);
        _cboTinh.value(record.maTinhpr_sd);
        if (record.ngungSD == "True")
            _chkNTD.checked(true);
        else
            _chkNTD.checked(false);
        Window1.setTitle("Sửa Quận/Huyện");
        setWindowPosition();
        return false;
    }
    function OnBeforeClientDelete(record)
    {//gan phan quyền
            var permiss = NTSoftReceipt.danhmuc.dmquanhuyen.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác xóa!");
                return false;
            }
        ma = record.maHuyenpr;
        if(NTSoftReceipt.danhmuc.dmquanhuyen.KTraXoa(ma).value)
        {
           alert("Quận/Huyện: " + record.tenHuyen + " đã được sử dụng nên bạn không thể thực hiện xóa!");
        }
        else
        {
            Dialog1.Open();
        }
        return false;
    }
    function LuuvaDongCN()
    {
        var _arrT = new Array();
            _arrT[0]=_txtMa.value();
            _arrT[1]=_txtTen.value();
            _arrT[2]=_txtGhiChu.value();
            _arrT[3]=_chkNTD.checked();
            _arrT[4]=_txtTenVT.value();
            _arrT[5]=_cboTinh.value();
       if(isEmptyText(_arrT[0])==true)
       {
            alert("Mã Quận/Huyện không được bỏ trống!");
            return false;
       }
       if(isEmptyText(_arrT[1])==true)
       {
            alert("Tên Quận/Huyện không được bỏ trống!");
            return false;
       }  
       if(isEmptyText(_arrT[5])==true)
       {
            alert("Tỉnh/Thành phố không được bỏ trống!");
            return false;
       }         
       if(coThaotac==1)
        {
            var kiemTra=NTSoftReceipt.danhmuc.dmquanhuyen.KTraTonTai(_arrT[0]);
            if(kiemTra.value==false)
            {
                NTSoftReceipt.danhmuc.dmquanhuyen.ThemDuLieu(_arrT); 
            }
            else
            {
               alert("Mã Quận/Huyện này đã tồn tại!");
               return false;
            }                
        }
        else
        {
            NTSoftReceipt.danhmuc.dmquanhuyen.SuaDuLieu(_arrT); 
        }
        Grid1.refresh();   
        Window1.Close();
        return false;
    }
    function XoaDoiTuong()
    {
        NTSoftReceipt.danhmuc.dmquanhuyen.XoaDuLieu(ma); 
        Grid1.refresh();   
        Dialog1.Close(); 
        return false;
    }
    </script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <cc1:Grid ID="Grid1" runat="server" FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true"
        PageSizeOptions="5,10,15,20" PageSize="15" AutoGenerateColumns="false"
        AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Height="550"
        OnRebind="Grid1_OnRebind" AllowAddingRecords="true" AllowMultiRecordSelection="False">
        <PagingSettings Position="Bottom" />
        <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
        <%--<ScrollingSettings ScrollHeight="100%" EnableVirtualScrolling="true" ScrollWidth="100%" />--%>
        <ClientSideEvents OnBeforeClientAdd="OnBeforeClientAdd" OnBeforeClientEdit="OnBeforeClientEdit"
            OnBeforeClientDelete="OnBeforeClientDelete" />
        <Columns>
            <cc1:Column AllowDelete="true" AllowEdit="true" HeaderText="Sửa/Xóa" Width="100px">
            </cc1:Column>
            <cc1:Column HeaderText="Mã" DataField="maHuyenpr" Width="120px" ReadOnly="true" Wrap ="true"
                Visible="true">
            </cc1:Column>
            <cc1:Column HeaderText="Tên Quận/Huyện" DataField="tenHuyen" Width="300px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Tên viết tắt" DataField="tenVietTat" Width="130px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="maTinhpr_sd" DataField="maTinhpr_sd" Width="200px" Wrap ="true" Visible="false">
            </cc1:Column>
            <cc1:Column HeaderText="Tỉnh/Thành phố" DataField="tenTinh" Width="250px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Ghi chú" DataField="ghiChu" Width="130px" Wrap="true" Visible="false">
            </cc1:Column>
            <cc1:CheckBoxColumn HeaderText="Ngưng theo dõi" DataField="ngungSD" Width="120px">
            </cc1:CheckBoxColumn>
        </Columns>
        <LocalizationSettings CancelAllLink="Hủy tất cả" AddLink="Thêm mới" CancelLink="Hủy"
            DeleteLink="Xóa" EditLink="Sửa" Filter_ApplyLink="Tìm kiếm" Filter_HideLink="Đóng tìm kiếm"
            Filter_RemoveLink="Xóa tìm kiếm" Filter_ShowLink="Mở tìm kiếm" 
            FilterCriteria_NoFilter="Không tìm kiếm"
                            FilterCriteria_Contains="Chứa"
                            FilterCriteria_DoesNotContain="Không chứa"
                            FilterCriteria_StartsWith="Bắt đầu với"
                            FilterCriteria_EndsWith="Kết thúc với"
                            FilterCriteria_EqualTo="Bằng"
                            FilterCriteria_NotEqualTo="Không bằng"
                            FilterCriteria_SmallerThan="Nhỏ hơn"
                            FilterCriteria_GreaterThan="Lớn hơn"
                            FilterCriteria_SmallerThanOrEqualTo="Nhỏ hơn hoặc bằng"
                            FilterCriteria_GreaterThanOrEqualTo="Lớn hơn hoặc bằng"
                            FilterCriteria_IsNull="Rỗng"
                            FilterCriteria_IsNotNull="Không rỗng"
                            FilterCriteria_IsEmpty="Trống"
                            FilterCriteria_IsNotEmpty="Không trống"
                            Paging_OfText="của"
            Grouping_GroupingAreaText="Kéo tiêu đề cột vào đây để loại theo cột đó" JSWarning="Có một lỗi khởi tạo lưới với ID '[GRID_ID]'. \ N \ n [Chú ý] \ n \ nHãy liên hệ bộ phận bảo trì của Nhất Tâm Soft để được giúp đỡ."
            LoadingText="Đang tải...." MaxLengthValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX vượt quá số lượng tối đa ký tự YYYYY cho phép cột này."
            ModifyLink="Chỉnh sửa" NoRecordsText="Không có dữ liệu" Paging_ManualPagingLink="Trang kế »"
            Paging_PageSizeText="Số dòng 1 trang:" Paging_PagesText="Trang:" Paging_RecordsText="Dòng:"
            ResizingTooltipWidth="Rộng:" SaveAllLink="Lưu tất cả" SaveLink="Lưu" TypeValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX là không đúng."
            UndeleteLink="Không xóa" UpdateLink="Lưu" />
        <TemplateSettings HeadingTemplateId="idHeadingGrid1" />
        <Templates>
            <cc1:GridTemplate runat="server" ID="idHeadingGrid1">
                <Template>
                    <%--<b style="color: #DA102D; font-style: italic;">Danh mục tình trạng hợp đồng</b>--%>
                    <b>Quận/Huyện</b>
                </Template>
            </cc1:GridTemplate>
        </Templates>
    </cc1:Grid>
    <owd:Window ID="Window1" runat="server" IsModal="true" ShowCloseButton="true" Status=""
        Title="Nhập báo cáo" ShowStatusBar="false" RelativeElementID="WindowPositionHelper"
        Height="260" Width="460" StyleFolder="~/App_Themes/Styles/wdstyles/dogma"
        VisibleOnLoad="false" IsResizable="false">
        <table border="0">
                        <tr>
                            <td>
                                <cc2:OboutButton ID="bttLuuvaDongCN" runat="server" Text="Lưu và đóng" Width="100px"
                                    FolderStyle="App_Themes/Styles/Interface/OboutButton" OnClientClick="LuuvaDongCN(); return false;">
                                </cc2:OboutButton>
                            </td>
                            <td>
                                <cc2:OboutButton ID="bttDongCN" runat="server" Text="Đóng" Width="100px" OnClientClick="Window1.Close(); return false;">
                                </cc2:OboutButton>
                            </td>
                            <td style="width: 200px" align="right">
                                <cc2:OboutCheckBox ID="_chkNTD" runat="server" Text="Ngưng theo dõi" FolderStyle="App_Themes/Styles/Interface/OboutCheckBox">
                                </cc2:OboutCheckBox>
                            </td>
                        </tr>
                    </table>
                    <div>
                        <fieldset style="border: 1px solid #DBDBE1">
            <legend>Thông tin Quận/Huyện</legend>
                            <table border="0">
                                <tr>
                                    <td style="width:100px">
                                        Mã 
                                    </td>
                                    <td style="width: 150px">
                                        <cc2:OboutTextBox ID="_txtMa" runat="server" Width="100%" FolderStyle="App_Themes/Styles/Interface/OboutTextBox"></cc2:OboutTextBox>
                                    </td>
                                    <td style="width: 150px">
                                    <span style="color: Red">*</span>
                                    </td>
                                    <td>
                                    
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Tên Quận/Huyện
                                    </td>
                                    <td colspan="2">
                                        <cc2:OboutTextBox ID="_txtTen" runat="server" Width="100%"></cc2:OboutTextBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Tên viết tắt
                                    </td>
                                    <td colspan="2">
                                        <cc2:OboutTextBox ID="_txtTenVT" runat="server" Width="100%"></cc2:OboutTextBox>
                                    </td>
                                    <td>
                                    <span style="color: Red"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Tỉnh/TP
                                    </td>
                                    <td colspan="2">
                                        <cc3:ComboBox ID="_cboTinh" runat="server" FilterType="Contains" Height="150px" Width="100%" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                        </cc3:ComboBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Ghi chú
                                    </td>
                                    <td colspan="2">
                                        <cc2:OboutTextBox ID="_txtGhiChu" runat="server" Width="100%"
                                            FolderStyle="App_Themes/Styles/Interface/OboutTextBox"></cc2:OboutTextBox>
                                    </td>
                                </tr>
                            </table>
                        </fieldset>
                </div>
    </owd:Window>
    <owd:Dialog ID="Dialog1" runat="server" IsModal="true" ShowCloseButton="true" Top="0" Left="250"
            Height="150" Width="350" VisibleOnLoad="false" StyleFolder="~/App_Themes/Styles/wdstyles/dogma"
            Title="Cảnh báo">
            <div align="center" style="height: 50px; margin-top:15px">
                Bạn có thật sự muốn xóa dòng dữ liệu đã chọn không? Đồng ý xóa chọn 'Thực hiện', không đồng ý chọn 'Bỏ qua'.
            </div>
            <div align="center">
                <table>
                    <tr>
                        <td>
                            <cc2:OboutButton ID="btnXoa" runat="server" Text="Thực hiện" Width="90px" OnClientClick="XoaDoiTuong(); return false;" FolderStyle="App_Themes/Styles/Interface/OboutButton">
                            </cc2:OboutButton>
                        </td>
                        <td>
                            <cc2:OboutButton ID="btnHuy" runat="server" Text="Bỏ qua" Width="90px" OnClientClick="Dialog1.Close(); return false;">
                            </cc2:OboutButton>
                        </td>
                    </tr>
                </table>
            </div>
        </owd:Dialog>
</asp:Content>
