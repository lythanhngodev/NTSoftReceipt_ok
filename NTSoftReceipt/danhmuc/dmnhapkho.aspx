<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Default.Master" CodeBehind="dmnhapkho.aspx.cs" Inherits="NTSoftReceipt.danhmuc.dmnhapkho" %>
<%@ Register Assembly="obout_Window_NET" Namespace="OboutInc.Window" TagPrefix="owd" %>
<%@ Register Assembly="obout_Grid_NET" Namespace="Obout.Grid" TagPrefix="cc1" %>
<%@ Register Assembly="obout_Interface" Namespace="Obout.Interface" TagPrefix="cc2" %>
<%@ Register Assembly="obout_ComboBox" Namespace="Obout.ComboBox" TagPrefix="cc3" %>
<%@ Register Assembly="obout_Calendar2_Net" Namespace="OboutInc.Calendar2" TagPrefix="obout" %>

<asp:Content runat="server" ContentPlaceHolderID="head">
    <script src="../js/jquery-3.3.1.min.js"></script>
    <script type="text/javascript">
        window.onload = function() {
            Grid1.refresh();
            Grid2.refresh();
        }
        function setWindowPositionW1() {
            var screenWidth = screen.width;
            var screenHeight = screen.height;
            var Window1Size = Window1.getSize();
            Window1.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(Window1Size.width)) / 2), 200);
            Window1.Open();
        }
        function setWindowPositionW2() {
            var screenWidth = screen.width;
            var screenHeight = screen.height;
            var Window2Size = Window2.getSize();
            Window2.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(Window2Size.width)) / 2), 200);
            Window2.Open();
        }
        var thaotacW2 = "them";
        var thaotacCT = "them";
        function LamMoi() {
            txtSoPhieu.value('');
            document.getElementById('txtNgayNhap').value = null;
            //cbNguoiNhap.value();
            //cbNhaCC.value();
            txtNoiDung.value('');
            txtGhiChu.value('');
        }
        function LamMoiNKCT(){
            txtSoLuongW2.value("");
            txtGiaNhapW2.value("");
            txtThanhTienW2.value("");
        }
        // làm tới đây
        function ThemMoiNhapKho() {
            Window1.setTitle("Thêm nhập kho");
            document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "lammoi";
            document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value = "0";
            thaotacW2 = "them";
            txtSoPhieu.enable();
            LamMoi();
            Window1.Open();
            Grid2.refresh();
            return false;
        }
        function OnBeforeClientEdit(record) {
            thaotacW2 = "sua";
            txtSoPhieu.value(record.soPhieu);
            txtSoPhieu.disable();
            //txtNgayNhap.value(record.ngayNhap);
            var date = new Date();
            try {
                
                date = record.ngayNhap.split("/");
                cbNguoiNhap.value(record.maNhanVienpr_sd);
                cbNhaCC.value(record.maNhaCCpr);
                txtNoiDung.value(record.noiDung);
                txtGhiChu.value(record.ghiChu);
                // Load Grid
                document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "loadgridtuSttNK";
                document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value = record.sttNKpr;
                Grid2.refresh();
                Window1.setTitle("Sửa thông tin nhập kho");
                setWindowPositionW1();
                $('#txtNgayNhap').val(date[2].substr(0,4)+"-"+date[1]+"-"+date[0]);
            } catch (e) {
                document.getElementById('txtNgayNhap').value = null;
            }
            return false;
        }
        function OnBeforeClientDelete(record){
            var cf = confirm('Bạn có muốn xóa nhập kho này?');
            if (cf) {
                var result = NTSoftReceipt.danhmuc.dmnhapkho.XoaNhapKho(record.sttNKpr).value;
                if (result) {
                    document.getElementById('<%= txtHDLoaiG1.ClientID%>').value = "lammoi";
                    Grid1.refresh();
                }else{
                    alert('Mục nhập kho này đang được sử dụng, không được xóa');
                }
            }
            return false;
        }
        function OnBeforeClientEditG2(record) {
            // chưa lưu
            if (!bttLuuCapNhatKho.IsDisabled) {
                alert('Vui lòng lưu nhập kho trước khi nhập kho chi tiết');
                return false;
            }
            else {
                cbBienLaiW2.value(record.sttBienLaipr_sd);
                txtSoLuongW2.value(record.soLuong);
                txtGiaNhapW2.value(record.donGia);
                txtThanhTienW2.value(record.thanhTien);
                document.getElementById('<%= txtHDsttNKCTpr.ClientID %>').value = record.sttNKCTpr;

                Window1.Close();
                setWindowPositionW2();
                bttLuuCapNhatKho.enable();
            }
            return false;
        }
        function OnBeforeClientDeleteG2(record){
            var cf = confirm('Bạn có muốn xóa chi tiết nhập kho này?');
            if (cf) {
                var result = NTSoftReceipt.danhmuc.dmnhapkho.XoaNhapKhoCT(record.sttNKCTpr).value;
                if (result) {
                    document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "loadgridtuSttNK";
                    Grid2.refresh();
                }
            }
            return false;
        }
        function LuuSuaNhapKho() {
            if (thaotacW2=="them") {
                var arrT = new Array();
                arrT[0] = txtSoPhieu.value();
                arrT[1] = document.getElementById('txtNgayNhap').value;
                arrT[2] = cbNguoiNhap.value();
                arrT[3] = cbNhaCC.value();
                arrT[4] = txtNoiDung.value();
                arrT[5] = txtGhiChu.value();
                for (var i = arrT.length - 1; i >= 0; i--) {
                    if (arrT[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                var result = NTSoftReceipt.danhmuc.dmnhapkho.ThemNhapKho(arrT).value;
                if (parseInt(result)!=0) {
                    document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "loadgridtuSttNK";
                    document.getElementById('<%= txtHDsttNKCTpr.ClientID %>').value = parseInt(result);
                    thaotacW2=="sua";
                    Grid2.refresh();
                    return false;
                } else {
                    alert('Quá trình thêm bị lỗi');
                }
                return false;
            }else if(thaotacW2=="sua"){
                var arrT = new Array();
                arrT[0] = document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value;
                arrT[1] = document.getElementById('txtNgayNhap').value;
                arrT[2] = cbNguoiNhap.value();
                arrT[3] = cbNhaCC.value();
                arrT[4] = txtNoiDung.value();
                arrT[5] = txtGhiChu.value();
                for (var i = arrT.length - 1; i >= 0; i--) {
                    if (arrT[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                var result = NTSoftReceipt.danhmuc.dmnhapkho.SuaNhapKho(arrT).value;
                if (result) {
                    bttLuuCapNhatKho.disable();
                    document.getElementById('<%= txtHDLoaiG1.ClientID%>').value = "lammoi";
                    Grid1.refresh();
                } else {
                    alert('Quá trính cập nhật bị lỗi');
                }
            }
            return false;
        }
        function LuuCapNhatNhapKhoCT() {
            if (thaotacCT=="them") {
                var arrT = new Array();
                arrT[0] = cbBienLaiW2.value();
                arrT[1] = txtSoLuongW2.value();
                arrT[2] = txtGiaNhapW2.value();
                arrT[3] = txtThanhTienW2.value();
                arrT[4] = document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value;
                for (var i = arrT.length - 1; i >= 0; i--) {
                    if (arrT[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                var result = NTSoftReceipt.danhmuc.dmnhapkho.ThemNhapKhoCT(arrT).value;
                if (result) {
                    Window1.Open();
                    Window2.Close();
                    document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "loadgridtuSttNK";
                    Grid2.refresh();
                    return false;
                } else {
                    alert('Quá trình cập nhật bị lỗi');
                }
                return false;
            }else if (thaotacCT=="sua"){
                var arrT = new Array();
                arrT[0] = cbBienLaiW2.value();
                arrT[1] = txtSoLuongW2.value();
                arrT[2] = txtGiaNhapW2.value();
                arrT[3] = txtThanhTienW2.value();
                arrT[4] = document.getElementById('<%= txtHDsttNKCTpr.ClientID%>').value;
                for (var i = arrT.length - 1; i >= 0; i--) {
                    if (arrT[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                var result = NTSoftReceipt.danhmuc.dmnhapkho.LuuCapNhatNhapKhoCT(arrT).value;
                if (result) {
                    Window1.Open();
                    Window2.Close();
                    document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "loadgridtuSttNK";
                    Grid2.refresh();
                    return false;
                } else {
                    alert('Quá trình cập nhật bị lỗi');
                }
            }
            return false;
        }
        function ThemMoiNKCT(){
            if (document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value.toString()!="0") {
                if (bttLuuCapNhatKho.IsDisabled) {
                    // cho thêm
                    thaotacCT = "them";
                    Window2.setTitle("Thêm mới NKCT");
                    LamMoiNKCT();
                    Window1.Close();
                    setWindowPositionW2();
                }else{
                    alert('Vui lòng ấn lưu trước');
                }
            }else{
                alert('Vui lòng ấn lưu trước');
            }
            return false;
        }
        function TinhThanhTien() {
            try {
                var soluong = parseInt(txtSoLuongW2.value());
                var gianhap = parseInt(txtGiaNhapW2.value());
                if (isNaN((soluong*gianhap).toString())) {
                    txtThanhTienW2.value("0");
                }else{
                    txtThanhTienW2.value((soluong*gianhap).toString());
                }
                
            } catch (e) {
                txtThanhTienW2.value("0");
            }
            return false;
        }
    </script>
</asp:Content>
<asp:Content runat="server" ContentPlaceHolderID="ContentPlaceHolder1">
    <asp:HiddenField runat="server" ID="txtHDLoaiG1" Value="lammoi" />
    <asp:HiddenField runat="server" ID="txtHDLoaiG2" Value="lammoi" />
    <asp:HiddenField runat="server" ID="txtSttNKpr_sd" Value="0" />
    <div><h2>Nhập kho</h2></div>
    <div style="width:100%">
        <button class="nut" onclick="ThemMoiNhapKho();return false;">Thêm</button>
    </div>
    <cc1:Grid ID="Grid1" runat="server" FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true"
        PageSizeOptions="5,10,15,20" PageSize="15" AutoGenerateColumns="false"
        AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Height="550" Width="500"
        OnRebind="Grid1_Rebind1" AllowAddingRecords="true" AllowMultiRecordSelection="False">
        <PagingSettings Position="Bottom" />
        <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
        <%--<ScrollingSettings ScrollHeight="100%" EnableVirtualScrolling="false" ScrollWidth="100%" />--%>
        <ClientSideEvents OnBeforeClientEdit="OnBeforeClientEdit" OnBeforeClientDelete="OnBeforeClientDelete" />
        <Columns>
            <cc1:Column AllowEdit="true" AllowDelete="true" HeaderText="Thao tác"></cc1:Column>
            <cc1:Column HeaderText="Mã TT" DataField="sttNKpr" Width="100px" ReadOnly="true" Visible="false">
            </cc1:Column>
            <cc1:Column HeaderText="Số phiếu" DataField="soPhieu" Width="100px" ReadOnly="true" Wrap ="true"
                Visible="true">
            </cc1:Column>
            <cc1:Column HeaderText="Ngày nhập" DataFormatString="{0:dd-MM-yyyy}" DataField="ngayNhap" Width="160" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Mã người lập" DataField="maNhanVienpr_sd" Width="100px" Visible="false">
            </cc1:Column>
            <cc1:Column HeaderText="Người lập" DataField="tenNhanVien" Width="200px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Mã nhà cung cấp" DataField="maNhaCCpr" Width="100px" Visible="false">
            </cc1:Column>
            <cc1:Column HeaderText="Nhà cung cấp" DataField="tenNhaCC" Width="200px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Nội dung" DataField="noiDung" Width="180px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Ghi chú" DataField="ghiChu" Width="150px" Wrap="true">
            </cc1:Column>
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
    </cc1:Grid>
    <owd:Window ID="Window1" runat="server" IsModal="true" ShowCloseButton="true" Status=""
        Title="Thêm / Sửa Phiếu nhập kho" ShowStatusBar="false" RelativeElementID="WindowPositionHelper"
        Height="600" Width="640" StyleFolder="~/App_Themes/Styles/wdstyles/dogma"
         IsResizable="true" VisibleOnLoad="false">
        <table border="0">
                        <tr>
                            <td>
                                <cc2:OboutButton ID="bttLuuCapNhatKho" runat="server" Text="Lưu" Width="100px"
                                    FolderStyle="App_Themes/Styles/Interface/OboutButton" OnClientClick="LuuSuaNhapKho(); return false;">
                                </cc2:OboutButton>
                            </td>
                            <td>
                                <cc2:OboutButton ID="bttDongCN" runat="server" Text="Đóng" Width="100px" OnClientClick="Window1.Close(); return false;">
                                </cc2:OboutButton>
                            </td>
                        </tr>
                    </table>
                    <div>
                        <fieldset style="border: 1px solid #DBDBE1">
                            <legend>Thông tin chung</legend>
                            <table border="0">
                                <tr>
                                    <td>
                                        Số phiếu 
                                    </td>
                                    <td style="width: 150px">
                                        <cc2:OboutTextBox ID="txtSoPhieu" runat="server" Width="100%" FolderStyle="App_Themes/Styles/Interface/OboutTextBox"></cc2:OboutTextBox>
                                    </td>
                                    <td>
                                        Ngày nhập
                                    </td>
                                    <td>
                                        <input type="date" value="" id="txtNgayNhap" />  
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Người nhập
                                    </td>
                                    <td>
                                        <cc3:ComboBox ID="cbNguoiNhap" runat="server" FilterType="Contains" EnableLoadOnDemand="true"  Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                        </cc3:ComboBox>
                                    </td>
                                    <td>
                                        Nhà cung cấp
                                    </td>
                                    <td>
                                        <cc3:ComboBox ID="cbNhaCC" runat="server" FilterType="Contains" EnableLoadOnDemand="true"  Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                        </cc3:ComboBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Nội dung
                                    </td>
                                    <td colspan="3">
                                        <cc2:OboutTextBox ID="txtNoiDung" runat="server" Width="100%"></cc2:OboutTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Ghi chú
                                    </td>
                                    <td colspan="3">
                                        <cc2:OboutTextBox ID="txtGhiChu" runat="server" Width="100%"
                                            FolderStyle="App_Themes/Styles/Interface/OboutTextBox"></cc2:OboutTextBox>
                                    </td>
                                </tr>
                            </table>
                        </fieldset>
                        <br />
                        <fieldset style="border: 1px solid #DBDBE1">
                            <legend>Thông tin chi tiết</legend>
                                <div class="head-grid-2">
                                    <input type="button" value="Thêm mới" onclick="ThemMoiNKCT()" />
                                    <input type="text" name="name" value="" style="float:right" placeholder="Thìm kiếm ..." />
                                </div>
                                <cc1:Grid ID="Grid2" runat="server" FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true"
                                    PageSizeOptions="5,10,15,20" PageSize="15" AutoGenerateColumns="false"
                                    AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Width="400" Height="300"
                                    OnRebind="Grid2_Rebind" AllowAddingRecords="true" AllowMultiRecordSelection="False">
                                    <PagingSettings Position="Bottom" />
                                    <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
                                    <%--<ScrollingSettings ScrollHeight="100%" EnableVirtualScrolling="false" ScrollWidth="100%" />--%>
                                    <ClientSideEvents OnBeforeClientEdit="OnBeforeClientEditG2" OnBeforeClientDelete="OnBeforeClientDeleteG2" />
                                    <Columns>
                                        <cc1:Column AllowEdit="true" AllowDelete="true" HeaderText="Thao tác" Width="100"></cc1:Column>
                                        <cc1:Column HeaderText="sttNKCTpr" DataField="sttNKCTpr" Visible="false"></cc1:Column>
                                        <cc1:Column HeaderText="Biên lai" DataField="sttBienLaipr_sd" Width="100" ReadOnly="true">
                                        </cc1:Column>
                                        <cc1:Column HeaderText="Số lượng" DataField="soLuong" Width="100" Wrap ="true">
                                        </cc1:Column>
                                        <cc1:Column HeaderText="Giá nhập" DataField="donGia" Width="130">
                                        </cc1:Column>
                                        <cc1:Column HeaderText="Thành tiền" DataField="thanhTien" Width="180" Wrap ="true">
                                        </cc1:Column>
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
                                </cc1:Grid>
                        </fieldset>
                </div>
    </owd:Window>
    <owd:Window ID="Window2" runat="server" IsModal="true" Status="" ShowCloseButton="false"
        Title="Sửa chi tiết nhập kho" ShowStatusBar="false" RelativeElementID="WindowPositionHelper"
        Height="160" Width="460" StyleFolder="~/App_Themes/Styles/wdstyles/dogma"
        VisibleOnLoad="false" IsResizable="false" >
        <table border="0">
            <tr>
                <td>
                    <cc2:OboutButton ID="bttLuuCT" runat="server" Text="Lưu và đóng" Width="100px"
                        FolderStyle="App_Themes/Styles/Interface/OboutButton" OnClientClick="LuuCapNhatNhapKhoCT(); return false;">
                    </cc2:OboutButton>
                </td>
                <td>
                    <cc2:OboutButton ID="bttDong" runat="server" Text="Đóng" Width="100px" OnClientClick="Window1.Open();Window2.Close(); return false;">
                    </cc2:OboutButton>
                </td>
                <td></td>
                <td></td>
            </tr>
        </table>
        <div>
            <fieldset style="border: 1px solid #DBDBE1">
                <legend>Sửa chi tiết nhập kho</legend>
                <table border="0">
                    <tr>
                        <td>
                            Biên lai
                        </td>
                        <td>
                            <asp:HiddenField runat="server" ID="txtHDsttNKCTpr" />
                            <cc3:ComboBox ID="cbBienLaiW2" runat="server" FilterType="Contains" Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                <ClientSideEvents  />
                            </cc3:ComboBox>
                        </td>
                        <td>
                            Số lượng
                        </td>
                        <td>
                            <cc2:OboutTextBox ID="txtSoLuongW2" runat="server"
                                FolderStyle="App_Themes/Styles/Interface/OboutTextBox">
                                <ClientSideEvents OnKeyUp="TinhThanhTien" />
                            </cc2:OboutTextBox>
                            </cc3:ComboBox>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            Giá nhập
                        </td>
                        <td>
                            <cc2:OboutTextBox ID="txtGiaNhapW2" runat="server"
                                FolderStyle="App_Themes/Styles/Interface/OboutTextBox">
                                <ClientSideEvents OnKeyUp="TinhThanhTien" />
                            </cc2:OboutTextBox>
                        </td>
                        <td>
                            Thành tiền
                        </td>
                        <td>
                            <cc2:OboutTextBox ID="txtThanhTienW2" runat="server"
                                FolderStyle="App_Themes/Styles/Interface/OboutTextBox" ReadOnly="true"></cc2:OboutTextBox>
                        </td>
                    </tr>
                </table>
            </fieldset>
    </div>
    </owd:Window>
</asp:Content>