<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Default.Master" CodeBehind="dmnhanvien.aspx.cs" Inherits="NTSoftReceipt.danhmuc.dmnhanvien" %>

<%@ Register Assembly="obout_Window_NET" Namespace="OboutInc.Window" TagPrefix="owd" %>
<%@ Register Assembly="obout_Grid_NET" Namespace="Obout.Grid" TagPrefix="cc1" %>
<%@ Register Assembly="obout_Interface" Namespace="Obout.Interface" TagPrefix="cc2" %>
<%@ Register Assembly="obout_ComboBox" Namespace="Obout.ComboBox" TagPrefix="cc3" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>

<asp:Content runat="server" ID="Content1" ContentPlaceHolderID="head">
    <style>
        .otim{
            float:right;
            margin-top:3px;
        }
        .nutthem{
            margin-top:3px;
        }
    </style>
    <script type="text/javascript">
        var thaotac = "them";
        var ma = "0";
        window.onload = function () {
            Grid1.refresh();
        }
        function setWindowPosition() {
            var screenWidth = screen.width;
            var screenHeight = screen.height;
            var Window1Size = Window1.getSize();
            Window1.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(Window1Size.width)) / 2), 200);
            Window1.Open();
        }
        function setWindowPositionW3() {
            var screenWidth = screen.width;
            var screenHeight = screen.height;
            var Window3Size = Window3.getSize();
            Window3.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(Window3Size.width)) / 2), 200);
            Window3.Open();
        }
        function LamMoiWindow1() {
            _txtMaNhanVien.disable();
            _txtMaNhanVien.value('');
            _txtTenNhanVien.value('');
            _rdGioiTinhNam.checked(false);
            _rdGioiTinhNu.checked(false);
            _txtSoDienThoai.value('');
            _txtDiaChi.value('');
            _txtGhiChu.value('');
            _cbPhongBan.value('');
        }
        function ThemMoi() {
            LamMoiWindow1();
            setWindowPosition();
            Window1.setTitle("Thêm nhân viên");
            thaotac = "them";
            _rdGioiTinhNam.checked(true);
            _txtMaNhanVien.value(NTSoftReceipt.danhmuc.dmnhanvien.TaoMaNhanVienTuDongAjax().value);
            //return false;
        }
        function TimKiem(t) {
            document.getElementById('<%= _txtKey.ClientID %>').value = t;
            document.getElementById('<%= _txtLoadRefesh.ClientID %>').value = "timkiem";
            Grid1.refresh();
        }
        function Luu() {
            if (thaotac == "them") {
                document.getElementById('<%= _txtLoadRefesh.ClientID %>').value = "lammoi";
                var _arrT = new Array();
                _arrT[0] = _txtTenNhanVien.value();
                _arrT[1] = (_rdGioiTinhNam.checked()) ? "Nam" : "Nữ";
                _arrT[2] = _txtDiaChi.value();
                _arrT[3] = _txtSoDienThoai.value();
                _arrT[4] = _txtThuNhap.value();
                _arrT[5] = _cbPhongBan.value();
                _arrT[6] = _txtGhiChu.value();
                _arrT[7] = _txtMaNhanVien.value();
                if (_arrT[0].trim().length == 0) {
                    alert("Nhập tên nhên viên");
                    return false;
                }
                if (_arrT[5].trim().length == 0) {
                    alert("Chọn phòng ban");
                    return false;
                }
                if (_arrT[4].trim().length == 0 || isNaN(_arrT[4])) {
                    alert("Nhập thu nhập");
                    return false;
                }
                var kq = NTSoftReceipt.danhmuc.dmnhanvien.ThemDuLieu(_arrT).value;
                Grid1.refresh();
            } else if (thaotac == "sua") {
                document.getElementById('<%= _txtLoadRefesh.ClientID %>').value = "lammoi";
                var _arrT = new Array();
                _arrT[0] = _txtTenNhanVien.value();
                _arrT[1] = (_rdGioiTinhNam.checked()) ? "Nam" : "Nữ";
                _arrT[2] = _txtDiaChi.value();
                _arrT[3] = _txtSoDienThoai.value();
                _arrT[4] = _txtThuNhap.value();
                _arrT[5] = _cbPhongBan.value();
                _arrT[6] = _txtGhiChu.value();
                _arrT[7] = _txtMaNhanVien.value();
                
                if (_arrT[0].trim().length==0) {
                    alert("Nhập tên nhên viên");
                    return false;
                }
                if (_arrT[5].trim().length==0) {
                    alert("Chọn phòng ban");
                    return false;
                }
                if (_arrT[4].trim().length==0 || isNaN(_arrT[4])) {
                    alert("Nhập thu nhập");
                    return false;
                }
                var kq = NTSoftReceipt.danhmuc.dmnhanvien.SuaDuLieu(_arrT).value;
                console.log(kq);
                Grid1.refresh();
            }
            Window1.Close();
            return false;
        }
        function XoaNhanVien() {
            NTSoftReceipt.danhmuc.dmnhanvien.XoaNhanVien(ma);
            Grid1.refresh();
            Dialog1.Close();
            return false;
        }
        function OnBeforeClientEdit(record) {
            LamMoiWindow1();
            Window1.setTitle("Chỉnh sửa thông tin");
            thaotac = "sua";
            _txtMaNhanVien.disable();
            _txtMaNhanVien.value(record.maNhanVienpr);
            _txtTenNhanVien.value(record.tenNhanVien);
            if (record.gioiTinh=="Nam") {
                _rdGioiTinhNam.checked(true);
                _rdGioiTinhNu.checked(false);

            } else if (record.gioiTinh=="Nữ") {
                _rdGioiTinhNu.checked(true);
                _rdGioiTinhNam.checked(false);
            } else {
                _rdGioiTinhNu.checked(false);
                _rdGioiTinhNam.checked(true);
            }
            _txtSoDienThoai.value(record.soDienThoai);
            _txtDiaChi.value(record.diaChi);
            _txtThuNhap.value(record.thuNhap);
            _txtGhiChu.value(record.ghiChu);
            _cbPhongBan.value(record.sttPhongBanpr);
            setWindowPosition();
            return false;
        }
        function OnBeforeClientDelete(record) {
            ma = record.maNhanVienpr;
            Dialog1.Open();
            return false;
        }
        function DiaChi() {
            _txtDiaChi.value(_cbXa.text() + ", " + _cbHuyen.text() + ", "+_cbTinh.text());
        }
        function _cbTinh_SelectedIndexChanged() {
            var MaHuyen = this.value();
            var ketqua = NTSoftReceipt.danhmuc.dmnhanvien.LoadCBHuyen(MaHuyen);
            _cbHuyen.options.clear();
            _cbXa.options.clear();
            for (var i = 0; i < ketqua.value.Rows.length; i++) {
                var maH = ketqua.value.Rows[i].maHuyenpr;
                var tenH = ketqua.value.Rows[i].tenHuyen;
                _cbHuyen.options.add(tenH, maH, i);
            }
            DiaChi();
            return false;
        }
        function _cbHuyen_SelectedIndexChanged() {
            var MaHuyen = this.value();
            var ketqua = NTSoftReceipt.danhmuc.dmnhanvien.LoadCBXa(MaHuyen);
            _cbXa.options.clear();
            for (var i = 0; i < ketqua.value.Rows.length; i++) {
                var maH = ketqua.value.Rows[i].maXapr;
                var tenH = ketqua.value.Rows[i].tenXa;
                _cbXa.options.add(tenH, maH, i);
            }
            DiaChi();
            return false;
        }
        function _cbXa_SelectedIndexChanged() {
            DiaChi();
        }
        function onClientOpen_Dialog1() {
            document.getElementById("info").innerHTML = "";
            return false;
        }
        function NhapExcel() {
            dlgUpload.Open();
            return false;
        }
        function uploadExcelError(sender, args) {
            alert("Tải tập tin Excel mẫu thất bại!");
        }
        function uploadExcelStarted() {
            document.getElementById("info").innerHTML = "<br/>Đang tải tệp tin...";
            document.getElementById('linkLoichitiet').style.visibility = "hidden";
        }
        function uploadExcelComplete(sender, args) {
            document.getElementById("info").innerHTML = "<br/>...";
            var fileExtension = args.get_fileName();
            if (fileExtension.indexOf('.xls') != -1) {
                if (cbXoaDuLieuCu.checked()) {
                        var kq = (NTSoftReceipt.danhmuc.dmnhanvien.NhapDuLieuExcel_Xoa());
                        document.getElementById("loiChiTiet").innerHTML = kq.value[0];
                        if (kq.value[0].trim().length>0) {
                            document.getElementById('linkLoichitiet').style.visibility = "visible";
                        } else {
                            document.getElementById('linkLoichitiet').style.visibility = "hidden";
                        }
                        document.getElementById("info").innerHTML = "Tổng số dòng: "+kq.value[3]+", Thành công: " + kq.value[1] + ", Không thành công: " + kq.value[2];
                        document.getElementById('info').style.visibility = "visible";
                    }
                else {
                        var kq = (NTSoftReceipt.danhmuc.dmnhanvien.NhapDuLieuExcel());
                        document.getElementById("loiChiTiet").innerHTML = kq.value[0];
                        if (kq.value[0].trim().length>0) {
                            document.getElementById('linkLoichitiet').style.visibility = "visible";
                            
                        } else {
                            document.getElementById('linkLoichitiet').style.visibility = "hidden";
                        }
                        document.getElementById("info").innerHTML = "Tổng số dòng: " + kq.value[3] + ", Thành công: " + kq.value[1] + ", Không thành công: " + kq.value[2];
                        document.getElementById('info').style.visibility = "visible";
                }
                document.getElementById('<%= _txtLoadRefesh.ClientID %>').value = "lammoi";
                Grid1.refresh();
            } else {
                document.getElementById("info").innerHTML = "<br/>Tập tin Excel mẫu không đúng định dạng.";
                return;
            }
        }
        function XemChiTietLoiNhapExcel() {
            DialogLoi.Open();
            return false;
        }
    </script>

</asp:Content>

<asp:Content runat="server" ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1">
    <h1>NHÂN VIÊN</h1>
    <asp:HiddenField runat="server" ID="_txtKey" />
    <asp:HiddenField runat="server" ID="_txtLoadRefesh" Value="lammoi" />
    <cc1:Grid ID="Grid1" runat="server" FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true"
        PageSizeOptions="15,20" PageSize="15" AutoGenerateColumns="false"
        AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Width="800" Height="500" AllowColumnResizing="true" 
        AllowAddingRecords="false" AllowMultiRecordSelection="False" AllowMultiRecordAdding="false" OnRebind="Grid1_Rebind" >
        <PagingSettings Position="Bottom" />
        <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
        <ScrollingSettings EnableVirtualScrolling="false" />
        <ClientSideEvents OnBeforeClientEdit="OnBeforeClientEdit" OnBeforeClientDelete="OnBeforeClientDelete" />
        <Columns>
            <cc1:Column AllowDelete="true" AllowEdit="true" HeaderText="Sửa/Xóa" Width="100px">
            </cc1:Column>
            <cc1:Column HeaderText="Mã NV" DataField="maNhanVienpr" Width="150px" ReadOnly="true" Wrap="true"
                Visible="true">
            </cc1:Column>
            <cc1:Column HeaderText="Tên NV" DataField="tenNhanVien" Width="300px" Wrap="true">
            </cc1:Column>
            <cc1:Column HeaderText="Giới tính" DataField="gioiTinh" Width="150px" Wrap="true">
            </cc1:Column>
            <cc1:Column HeaderText="Địa chỉ" DataField="diaChi" Width="180px" Wrap="true">
            </cc1:Column>
            <cc1:Column HeaderText="SĐT" DataField="soDienThoai" Width="120px">
            </cc1:Column>
            <cc1:Column HeaderText="NS" DataField="namSinh" DataFormatString="{0:dd/MM/yyyy}" Width="120px">
            </cc1:Column>
            <cc1:Column HeaderText="Thu nhập" DataField="thuNhap" DataFormatString="{0:#,#0}" Width="150px" Wrap="true">
            </cc1:Column>
            <cc1:Column HeaderText="Phòng ban" DataField="tenPhongBan" Width="180px" Wrap="true">
            </cc1:Column>
            <cc1:Column HeaderText="Mã phong ban" DataField="sttPhongBanpr" Width="100px" Visible="false">
            </cc1:Column>
            <cc1:Column HeaderText="Ghi chú" DataField="ghiChu" Width="120px">
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
        <TemplateSettings HeadingTemplateId="idHeadingGrid1" />
        <Templates>
            <cc1:GridTemplate runat="server" ID="idHeadingGrid1">
                <Template>
                    <asp:Label>Nhân Viên</asp:Label>
                    <input type="button" value="Thêm mới" class="nutthem" onclick="ThemMoi()" />
                    <input type="button" value="Nhập từ Excel" class="nutthem" onclick="NhapExcel()" />
                    <input
                        type="text"
                        style="width: 250px"
                        placeholder="Nội dung tìm kiếm..."
                        class="otim" onkeyup="TimKiem(this.value)" />
                </Template>
            </cc1:GridTemplate>
        </Templates>
    </cc1:Grid>
    <owd:Window ID="Window1" runat="server" IsModal="true" ShowCloseButton="true" Status=""
        Title="Thêm nhân viên" ShowStatusBar="false" RelativeElementID="WindowPositionHelper"
        Height="360" Width="640" StyleFolder="~/App_Themes/Styles/wdstyles/dogma"
        VisibleOnLoad="false" IsResizable="false">
        <table border="0">
                        <tr>
                            <td>
                                <cc2:OboutButton ID="bttLuu" runat="server" Text="Lưu và đóng" Width="100px"
                                    FolderStyle="App_Themes/Styles/Interface/OboutButton" OnClientClick="Luu(); return false;">
                                </cc2:OboutButton>
                            </td>
                            <td>
                                <cc2:OboutButton ID="bttDong" runat="server" Text="Đóng" Width="100px" OnClientClick="Window1.Close(); return false;">
                                </cc2:OboutButton>
                            </td>
                            <td></td>
                            <td></td>
                        </tr>
                    </table>
                    <div>
                        <fieldset style="border: 1px solid #DBDBE1">
            <legend>Thông tin nhân viên</legend>
                            <table border="0">
                                <tr>
                                    <td style="width:100px">
                                        Mã NV 
                                    </td>
                                    <td style="width: 150px">
                                        <cc2:OboutTextBox ID="_txtMaNhanVien" runat="server" FolderStyle="App_Themes/Styles/Interface/OboutTextBox"></cc2:OboutTextBox>
                                    </td>
                                    <td style="width: 150px">
                                    <span style="color: Red">*</span>
                                    </td>
                                    <td style="width:100px;">
                                        Tên nhân viên
                                    </td>
                                    <td>
                                        <cc2:OboutTextBox ID="_txtTenNhanVien" runat="server"></cc2:OboutTextBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                </tr>

                                <tr>
                                    <td>
                                        Giới tính
                                    </td>
                                    <td colspan="2">
                                        <cc2:OboutRadioButton runat="server" ID="_rdGioiTinhNam" GroupName="_ggioitinh" Text="Nam">
                                        </cc2:OboutRadioButton>
                                        <cc2:OboutRadioButton runat="server" ID="_rdGioiTinhNu" GroupName="_ggioitinh" Text="Nữ">
                                        </cc2:OboutRadioButton>
                                    </td>
                                    <td>
                                        Số điện thoại
                                    </td>
                                    <td>
                                        <cc2:OboutTextBox ID="_txtSoDienThoai" runat="server"></cc2:OboutTextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Tỉnh/TP
                                    </td>
                                    <td>
                                        <cc3:ComboBox ID="_cbTinh" runat="server" FilterType="Contains" EnableLoadOnDemand="true"  Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                            <ClientSideEvents OnSelectedIndexChanged="_cbTinh_SelectedIndexChanged" />
                                        </cc3:ComboBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                    <td>
                                        Huyện
                                    </td>
                                    <td>
                                        <cc3:ComboBox ID="_cbHuyen" runat="server" FilterType="Contains" Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                            <ClientSideEvents OnSelectedIndexChanged="_cbHuyen_SelectedIndexChanged" />
                                        </cc3:ComboBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Xã
                                    </td>
                                    <td>
                                        <cc3:ComboBox ID="_cbXa" runat="server" FilterType="Contains" Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
                                            <ClientSideEvents OnSelectedIndexChanged="_cbXa_SelectedIndexChanged" />
                                        </cc3:ComboBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                    <td>
                                        Địa chỉ
                                    </td>
                                    <td>
                                        <cc2:OboutTextBox ID="_txtDiaChi" runat="server"
                                            FolderStyle="App_Themes/Styles/Interface/OboutTextBox" TextMode="MultiLine" ReadOnly="true"></cc2:OboutTextBox>
                                    </td>
                                    <td>
                                    <span style="color: Red">*</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        Thu nhập
                                    </td>
                                    <td colspan="2">
                                        <cc2:OboutTextBox ID="_txtThuNhap" runat="server"
                                            FolderStyle="App_Themes/Styles/Interface/OboutTextBox"></cc2:OboutTextBox>
                                    </td>
                                    <td>
                                        Phòng ban
                                    </td>
                                    <td colspan="2">
                                        <cc3:ComboBox ID="_cbPhongBan" runat="server" FilterType="Contains" Height="150px" FolderStyle="App_Themes/Styles/Interface/OboutComboBox">
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
                                    <td colspan="5">
                                        <cc2:OboutTextBox ID="_txtGhiChu" runat="server"
                                            FolderStyle="App_Themes/Styles/Interface/OboutTextBox" Width="100%"></cc2:OboutTextBox>
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
                Bạn có muốn xóa dòng dữ liệu đã chọn không? Đồng ý xóa chọn 'Thực hiện', không đồng ý chọn 'Bỏ qua'.
            </div>
            <div align="center">
                <table>
                    <tr>
                        <td>
                            <cc2:OboutButton ID="btnXoa" runat="server" Text="Thực hiện" Width="90px" OnClientClick="XoaNhanVien(); return false;" FolderStyle="App_Themes/Styles/Interface/OboutButton">
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
    <ajaxToolkit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
    </ajaxToolkit:ToolkitScriptManager>
    <owd:Dialog ID="dlgUpload" runat="server" IsModal="true" ShowCloseButton="true" Top="0"
        Left="250" Height="200" Width="400" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        Title="Chọn tập tin" OnClientOpen="onClientOpen_Dialog1()">
        <div style="padding-top: 10px">
            <cc2:OboutCheckBox ID="cbXoaDuLieuCu" runat="server" Text="Xóa dữ liệu cũ"
                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutCheckBox">
            </cc2:OboutCheckBox>
            <ajaxToolkit:AsyncFileUpload OnClientUploadError="uploadExcelError" OnClientUploadComplete="uploadExcelComplete"
                OnClientUploadStarted="uploadExcelStarted" runat="server" ID="AsyncFileUpload"
                Width="390px" UploaderStyle="Traditional" UploadingBackColor="" ThrobberID="myThrobber"
                BorderStyle="NotSet" Font-Underline="False" Font-Strikeout="False" />
            <br />
            <span><i>Chỉ cho phép tải lên tập tin định dạng: xls,xlsx</i></span><br />
            <span id="info" style="color: Red"></span><br />
            <a id="linkLoichitiet" style="visibility:hidden;cursor:pointer" onclick="XemChiTietLoiNhapExcel();return false;"><u>Xem chi tiết lỗi >></u></a>
        </div>
    </owd:Dialog>
    <owd:Dialog ID="DialogLoi" runat="server" IsModal="true" ShowCloseButton="true" Top="0"
        Left="250" Height="350" Width="400" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        Title="Chọn tập tin" OnClientOpen="onClientOpen_Dialog1()">
        <div style="padding-top: 10px;height:350px;overflow:auto;" id="loiChiTiet">
        </div>
    </owd:Dialog>

</asp:Content>