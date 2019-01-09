<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Default.Master" CodeBehind="dmnhapkhoNC.aspx.cs" Inherits="NTSoftReceipt.danhmuc.dmnhapkhoNC" %>
<%@ Register Assembly="obout_Grid_NET" Namespace="Obout.Grid" TagPrefix="cc1" %>
<asp:Content runat="server" ContentPlaceHolderID="head">
    <link href="../App_Themes/BuocThaoTac/css/bootstrap.min.css" rel="stylesheet" />
    <link href="../App_Themes/BuocThaoTac/css/ace.min.css" rel="stylesheet" />
    <link href="../App_Themes/BuocThaoTac/font-awesome/4.7.0/css/font-awesome.min.css" rel="stylesheet" />
    <link href="../Content/themes/Theme_1/css/select2.min.css" rel="stylesheet" />
    <script src="../App_Themes/BuocThaoTac/js/jquery-2.1.4.min.js"></script>
    <script src="../App_Themes/BuocThaoTac/js/ace-elements.min.js"></script>
    <script src="../App_Themes/BuocThaoTac/js/ace.min.js"></script>
</asp:Content>
<asp:Content runat="server" ContentPlaceHolderID="ContentPlaceHolder1">
    <asp:HiddenField runat="server" ID="txtHDLoaiG1" Value="lammoi" />
    <asp:HiddenField runat="server" ID="txtHDLoaiG2" Value="lammoi" />
    <asp:HiddenField runat="server" ID="txtSttNKpr_sd" Value="0" />
    <div class="main-content">
        <div class="row">
            <div class="col-md-12">
                <div class="col-md-9" id="khung1" style="display: block;">
                    <div class="widget-box">
                        <div class="widget-header">
                            <h4 class="widget-title">Nhập kho
                            </h4>
                        </div>

                        <div class="widget-body">
                            <div class="widget-main">
                                <div class="control-group">
                                    <div class="row">
                                        <div class="col-md-12">
                                            <input type="button" id="themmoi" class="btn btn-sm btn-primary" value="Thêm mới" />
                                            <hr />
                                        </div>
                                        <div class="col-xs-12">
                                            <div class="row" style="padding-bottom: 2px; padding-top: 5px">
                                                <div class="col-md-6">
                                                </div>
                                                <div class="col-md-6" style="float: right">
                                                    <input style="width: 200px; float: right;" type="text" placeholder="Nội dung tìm kiếm..." onkeyup="searchValue(Grid1,1,this.value)" class="searchCss">
                                                </div>
                                            </div>
                                        </div>
                                        <div class="col-md-12">
                                            <cc1:Grid ID="Grid1" runat="server" FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true" FilterType="ProgrammaticOnly"
                                                PageSizeOptions="5,10,15,20,-1" PageSize="15" AutoGenerateColumns="false" AllowFiltering="true" AllowAddingRecords="false"
                                                EnableRecordHover="true" AllowGrouping="false" Height="550" Width="100%" AllowMultiRecordSelection="false"
                                                OnRebind="Grid1_Rebind1">
                                                <ScrollingSettings ScrollWidth="100%" ScrollHeight="100%" />
                                                <PagingSettings Position="Bottom" />
                                                <FilteringSettings MatchingType="AnyFilter" />
                                                <ExportingSettings KeepColumnSettings="true" Encoding="UTF8" />
                                                <Columns>
                                                    <cc1:Column ParseHTML="true" ReadOnly="true" HeaderText="Thao tác" AllowSorting="false" Width="80px" AllowFilter="false" DataField="thaoTac">
                                                    </cc1:Column>
                                                    <cc1:Column HeaderText="Số phiếu" DataField="soPhieu">
                                                    </cc1:Column>
                                                    <cc1:Column HeaderText="Ngày nhập" DataFormatString="{0:dd-MM-yyyy}" DataField="ngayNhap">
                                                    </cc1:Column>
                                                    <cc1:Column HeaderText="Người lập" DataField="tenNhanVien">
                                                    </cc1:Column>
                                                    <cc1:Column HeaderText="Nhà cung cấp" DataField="tenNhaCC" >
                                                    </cc1:Column>
                                                    <cc1:Column HeaderText="Nội dung" DataField="noiDung">
                                                    </cc1:Column>
                                                    <cc1:Column HeaderText="Ghi chú" DataField="ghiChu">
                                                    </cc1:Column>
                                                </Columns>
                                                <LocalizationSettings CancelAllLink="Hủy tất cả" AddLink="Thêm mới" CancelLink="Hủy"
                                                    DeleteLink="Xóa" EditLink="Sửa" Filter_ApplyLink="Tìm kiếm" Filter_HideLink="Đóng tìm kiếm"
                                                    Filter_RemoveLink="Xóa tìm kiếm" Filter_ShowLink="Mở tìm kiếm" FilterCriteria_NoFilter="Không tìm kiếm"
                                                    FilterCriteria_Contains="Chứa" FilterCriteria_DoesNotContain="Không chứa" FilterCriteria_StartsWith="Bắt đầu với"
                                                    FilterCriteria_EndsWith="Kết thúc với" FilterCriteria_EqualTo="Bằng" FilterCriteria_NotEqualTo="Không bằng"
                                                    FilterCriteria_SmallerThan="Nhỏ hơn" FilterCriteria_GreaterThan="Lớn hơn" FilterCriteria_SmallerThanOrEqualTo="Nhỏ hơn hoặc bằng"
                                                    FilterCriteria_GreaterThanOrEqualTo="Lớn hơn hoặc bằng" FilterCriteria_IsNull="Rỗng"
                                                    FilterCriteria_IsNotNull="Không rỗng" FilterCriteria_IsEmpty="Trống" FilterCriteria_IsNotEmpty="Không trống"
                                                    Paging_OfText="của" Grouping_GroupingAreaText="Kéo tiêu đề cột vào đây để loại theo cột đó"
                                                    JSWarning="Có một lỗi khởi tạo lưới với ID '[GRID_ID]'. \ N \ n [Chú ý] \ n \ nHãy liên hệ bộ phận bảo trì của Nhất Tâm Soft để được giúp đỡ."
                                                    LoadingText="Đang tải...." MaxLengthValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX vượt quá số lượng tối đa ký tự YYYYY cho phép cột này."
                                                    ModifyLink="Chỉnh sửa" NoRecordsText="Không có dữ liệu" Paging_ManualPagingLink="Trang kế »"
                                                    Paging_PageSizeText="Số dòng 1 trang:" Paging_PagesText="Trang:" Paging_RecordsText="Dòng:"
                                                    ResizingTooltipWidth="Rộng:" SaveAllLink="Lưu tất cả" SaveLink="Lưu" TypeValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX là không đúng."
                                                    UndeleteLink="Không xóa" UpdateLink="Lưu" />
                                            </cc1:Grid>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="col-md-9" id="khung2" style="display: none;">
                    <div class="widget-box">
                        <div class="widget-header">
                            <h4 class="widget-title">Thêm mới nhập kho
                            </h4>
                        </div>

                        <div class="widget-body">
                            <div class="widget-main">
                                <div class="control-group">
                                    <div class="row">
                                        <div class="col-md-12">
                                            <div class="col-md-12">
                                                <button id="btnThemmoi" class="btn btn-sm btn-success">Thêm</button>&ensp;
                                                <button id="btnDong" class="btn btn-sm btn-danger">Đóng</button>
                                                <hr />
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Số phiếu</label>
                                                <div class="col-sm-9">
                                                    <input type="text" id="_sophieu" placeholder="Số phiếu ..." class="form-control">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Ngày nhập</label>
                                                <div class="col-sm-9">
                                                    <input type="date" id="_ngaynhap" class="form-control">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Người nhập</label>
                                                <div class="col-sm-9">
                                                    <select class="form-control" id="_nguoinhap">
                                                    </select>
                                                </div>
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Nhà cung cấp</label>
                                                <div class="col-sm-9">
                                                    <select class="form-control" id="_nhacungcap">
                                                    </select>
                                                </div>
                                            </div>
                                            <div class="form-group col-md-12">
                                                <label class="col-sm-2 control-label no-padding-right">Nội dung</label>
                                                <div class="col-sm-10">
                                                    <input type="text" class="form-control" id="_noidung">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-12">
                                                <label class="col-sm-2 control-label no-padding-right">Ghi chú</label>
                                                <div class="col-sm-10">
                                                    <input type="text" class="form-control" id="_ghichu">
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>

                <!-- khung 3 -->
                <div class="col-md-9" id="khung3" style="display: none;">
                    <div class="widget-box">
                        <div class="widget-header">
                            <h4 class="widget-title">Chỉnh sửa nhập kho
                            </h4>
                        </div>

                        <div class="widget-body">
                            <div class="widget-main">
                                <div class="control-group">
                                    <div class="row">
                                        <div class="col-md-12">
                                            <div class="col-md-12">
                                                <button id="btnSua" class="btn btn-sm btn-success">Sửa</button>&ensp;
                                                <button id="btnDongSua" class="btn btn-sm btn-danger">Đóng</button>
                                                <hr />
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Số phiếu</label>
                                                <div class="col-sm-9">
                                                    <input type="text" id="_ssophieu" placeholder="Số phiếu ..." class="form-control">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Ngày nhập</label>
                                                <div class="col-sm-9">
                                                    <input type="date" id="_sngaynhap" class="form-control">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Người nhập</label>
                                                <div class="col-sm-9">
                                                    <select class="form-control" id="_snguoinhap">
                                                    </select>
                                                </div>
                                            </div>
                                            <div class="form-group col-md-6">
                                                <label class="col-sm-3 control-label no-padding-right">Nhà cung cấp</label>
                                                <div class="col-sm-9">
                                                    <select class="form-control" id="_snhacungcap">
                                                    </select>
                                                </div>
                                            </div>
                                            <div class="form-group col-md-12">
                                                <label class="col-sm-2 control-label no-padding-right">Nội dung</label>
                                                <div class="col-sm-10">
                                                    <input type="text" class="form-control" id="_snoidung">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-12">
                                                <label class="col-sm-2 control-label no-padding-right">Ghi chú</label>
                                                <div class="col-sm-10">
                                                    <input type="text" class="form-control" id="_sghichu">
                                                </div>
                                            </div>
                                            <div class="form-group col-md-12">
                                                <button id="btnThemCTNK" class="btn btn-sm btn-success">Thêm mới chi tiết</button>
                                            </div>
                                        <div class="col-xs-12">
                                            <div class="row" style="padding-bottom: 2px; padding-top: 5px">
                                                <div class="col-md-6">
                                                </div>
                                                <div class="col-md-6" style="float: right">
                                                    <input style="width: 200px; float: right;" type="text" placeholder="Nội dung tìm kiếm..." onkeyup="searchValue(Grid2,1,this.value)" class="searchCss">
                                                </div>
                                            </div>
                                        </div>
                                            <div class="form-group col-md-12">
                                                <cc1:Grid ID="Grid2" runat="server" FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true"
                                                    PageSizeOptions="5,10,15,20" PageSize="15" AutoGenerateColumns="false" FilterType="ProgrammaticOnly"
                                                    AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Height="300"
                                                    OnRebind="Grid2_Rebind" AllowAddingRecords="false" AllowMultiRecordSelection="False">
                                                    <ScrollingSettings ScrollWidth="100%" ScrollHeight="100%" />
                                                    <PagingSettings Position="Bottom" />
                                                    <FilteringSettings MatchingType="AnyFilter" />
                                                    <ExportingSettings KeepColumnSettings="true" Encoding="UTF8" />
                                                    <Columns>
                                                        <cc1:Column ParseHTML="true" DataField="thaoTac" HeaderText="Thao tác" ></cc1:Column>
                                                        <cc1:Column HeaderText="sttNKCTpr" DataField="sttNKCTpr" Visible="false"></cc1:Column>
                                                        <cc1:Column HeaderText="Biên lai" DataField="tenBienLai" ReadOnly="true">
                                                        </cc1:Column>
                                                        <cc1:Column HeaderText="Số lượng" DataField="soLuong" DataFormatString="{0:#,#0}" Wrap="true">
                                                        </cc1:Column>
                                                        <cc1:Column HeaderText="Giá nhập" DataField="donGia" DataFormatString="{0:#,#0}">
                                                        </cc1:Column>
                                                        <cc1:Column HeaderText="Thành tiền" DataField="thanhTien" DataFormatString="{0:#,#0}" Wrap="true">
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
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div id="modalThemNKCT" class="modal" role="dialog">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
                    <h3 class="smaller lighter blue no-margin">Thêm chi tiết nhập kho</h3>
                </div>
                <div class="modal-body">
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Biên lai</label>
                            <div class="col-sm-9">
                                <select class="form-control clsbienlai" id="_ctbienlai">
                                </select>
                            </div>
                        </div>
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Số lượng</label>
                            <div class="col-sm-9">
                                <input type="number" onkeyup="TinhThanhTien();return false;" onchange="TinhThanhTien();return false;" class="form-control" id="_ctsoluong">
                            </div>
                        </div>
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Giá nhập</label>
                            <div class="col-sm-9">
                                <input type="number" onkeyup="TinhThanhTien();return false;" onchange="TinhThanhTien();return false;" class="form-control" id="_ctgianhap">
                            </div>
                        </div>
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Thành tiền</label>
                            <div class="col-sm-9">
                                <input type="number" class="form-control" id="_ctthanhtien" readonly="readonly">
                            </div>
                        </div>
                </div>
                <div class="modal-footer">
                    <div class="pull-right">
                        <button class="btn btn-sm btn-danger" data-dismiss="modal">Đóng</button>&ensp;
                        <button id="btnLuuThemCTNK" class="btn btn-sm btn-success">Thêm chi tiết</button>                     
                    </div>
                </div>
            </div>
            <!-- /.modal-content -->
        </div>
        <!-- /.modal-dialog -->
    </div>
    <div id="modalSuaNKCT" class="modal" role="dialog">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
                    <h3 class="smaller lighter blue no-margin">Sửa chi tiết nhập kho</h3>
                </div>
                <div class="modal-body">
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Biên lai</label>
                            <div class="col-sm-9">
                                <select class="form-control clsbienlai" id="_sctbienlai">
                                </select>
                            </div>
                        </div>
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Số lượng</label>
                            <div class="col-sm-9">
                                <input type="number" onkeyup="TinhThanhTienS();return false;" onchange="TinhThanhTienS();return false;" class="form-control" id="_sctsoluong">
                            </div>
                        </div>
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Giá nhập</label>
                            <div class="col-sm-9">
                                <input type="number" onkeyup="TinhThanhTienS();return false;" onchange="TinhThanhTienS();return false;" class="form-control" id="_sctgianhap">
                            </div>
                        </div>
                        <div class="form-group col-md-12">
                            <label class="col-sm-3 control-label no-padding-right">Thành tiền</label>
                            <div class="col-sm-9">
                                <input type="number" class="form-control" id="_sctthanhtien" readonly="readonly">
                            </div>
                        </div>
                        <input type="hidden" hidden="hidden" id="_sctsttNKCTpr">
                </div>
                <div class="modal-footer">
                    <div class="pull-right">
                        <button class="btn btn-sm btn-danger" data-dismiss="modal">Đóng</button>&ensp;
                        <button id="btnLuuSuaCTNK" class="btn btn-sm btn-success">Sửa chi tiết</button>                     
                    </div>
                </div>
            </div>
            <!-- /.modal-content -->
        </div>
        <!-- /.modal-dialog -->
    </div>
<style type="text/css">
    .row{
        padding:10px;
    }
    .ob_gMCont{
        width:100% !important;
        position: relative !important;
    }
</style>
    <script src="../Content/themes/Theme_1/js/bootstrap.min.js"></script>
    <script src="../Content/themes/Theme_1/js/jquery.dataTables.min.js"></script>
    <script src="../Content/themes/Theme_1/js/jquery.dataTables.bootstrap.min.js"></script>
    <script src="../Content/themes/Theme_1/material-preloader/js/materialPreloader.js"></script>
    <script src="../Content/themes/Theme_1/js/jquery.gritter.min.js"></script>
    <script src="../Content/themes/Theme_1/js/bootbox.js"></script>
    <script src="../Content/themes/Theme_1/js/select2.min.js"></script>
    <script src="../Content/themes/Theme_1/js/ace-elements.min.js"></script>
    <script src="../Content/themes/Theme_1/js/ace.min.js"></script>
    <script src="../Content/themes/Theme_1/js/dataTables.select.min.js"></script>
    <script src="../Content/themes/Theme_1/js/bootstrap-datepicker.min.js"></script>
    <script src="../js/NTSLibrary.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            $(document).on('click', '#themmoi', function () {
                $('#khung1').hide(234);
                $('#khung2').show(234);
                $('#khung2').find('input,select').val('');
            });
            $.ajax({
                url: 'dmnhapkhoNC.aspx/LayNguoiNhap',
                type: 'POST',
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    var a = $.parseJSON(data.d);
                    $('#_nguoinhap').empty();
                    $.each(a, function () {
                        $('#_nguoinhap').append($('<option></option>').val(this['maNhanVienpr']).html(this['tenNhanVien']));
                        $('#_snguoinhap').append($('<option></option>').val(this['maNhanVienpr']).html(this['tenNhanVien']));
                    });
                }
            });
            $.ajax({
                url: 'dmnhapkhoNC.aspx/LayNhaCungCap',
                type: 'POST',
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    var a = $.parseJSON(data.d);
                    $('#_nhacungcap').empty();
                    $('#_snhacungcap').empty();
                    $.each(a, function () {
                        $('#_nhacungcap').append($('<option></option>').val(this['maNhaCCpr']).html(this['tenNhaCC']));
                        $('#_snhacungcap').append($('<option></option>').val(this['maNhaCCpr']).html(this['tenNhaCC']));
                    });
                }
            });
            $.ajax({
                url: 'dmnhapkhoNC.aspx/LayBienLai',
                type: 'POST',
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    var a = $.parseJSON(data.d);
                    $('.clsbienlai').empty();
                    $.each(a, function () {
                        $('.clsbienlai').append($('<option></option>').val(this['sttBienLaipr']).html(this['tenBienLai']));
                    });
                }
            });
            $(document).on('click', '#btnDong', function () {
                $('#khung2').hide(234);
                $('#khung1').show(234);
                return false;
            });
            $(document).on('click', '#btnDongSua', function () {
                $('#khung1').show(234);
                $('#khung3').hide(234);
                $('#khung3').find('input,select').val('');
                Grid1.refresh();
                return false;
            });
            $(document).on('click', '#btnThemmoi', function () {
                var _sophieu = $('#_sophieu').val();
                var _ngaynhap = $('#_ngaynhap').val();
                var _nguoinhap = $('#_nguoinhap').val();
                var _nhacungcap = $('#_nhacungcap').val();
                var _noidung = $('#_noidung').val();
                var _ghichu = $('#_ghichu').val();
                if (jQuery.isEmptyObject(_sophieu)) {
                    alert('Vui lòng nhập số phiếu');
                    return false;
                }
                if (jQuery.isEmptyObject(_ngaynhap)) {
                    alert('Vui lòng chọn ngày nhập');
                    return false;
                }
                if (jQuery.isEmptyObject(_nguoinhap)) {
                    alert('Vui lòng chọn người nhập');
                    return false;
                }
                if (jQuery.isEmptyObject(_nhacungcap)) {
                    alert('Vui lòng chọn nhà cung cấp');
                    return false;
                }
                var obj = [];
                obj.push(_sophieu);
                obj.push(_ngaynhap);
                obj.push(_nguoinhap);
                obj.push(_nhacungcap);
                obj.push(_noidung);
                obj.push(_ghichu);
                $.ajax({
                    url: 'dmnhapkhoNC.aspx/ThemNhapKho',
                    type: 'POST',
                    data: "{obj:" + JSON.stringify(obj) + "}",
                    contentType: "application/json; charset=utf-8",
                    success: function (data) {
                        if ($.parseJSON(data.d)) {
                            alert('Thêm thành công');
                            $('#khung2').hide(234);
                            $('#khung1').show(234);
                            Grid1.refresh();
                        }
                        else {
                            alert('Thêm không thành công, kiểm tra lại thông tin');
                        }
                    }
                });
                return false;
            });
            $(document).on('click','#btnSua',function(){
                var obj = new Array();
                obj[0] = document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value;
                obj[1] = $('#_sngaynhap').val();
                obj[2] = $('#_snguoinhap').val();
                obj[3] = $('#_snhacungcap').val();
                obj[4] = $('#_snoidung').val();
                obj[5] = $('#_sghichu').val();
                for (var i = obj.length - 1; i >= 0; i--) {
                    if (obj[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                $.ajax({
                    url: 'dmnhapkhoNC.aspx/SuaNhapKho',
                    type: 'POST',
                    data: "{obj:" + JSON.stringify(obj) + "}",
                    contentType: "application/json; charset=utf-8",
                    success: function (data) {
                        if ($.parseJSON(data.d)) {
                            alert('Sửa thành công');
                            $('#btnSua').attr('disabled', 'disabled');
                        }
                        else {
                            alert('Thêm không thành công, kiểm tra lại thông tin');
                        }
                    }
                });
                return false;
            });
            $('#_nguoinhap, #_nhacungcap,#_snguoinhap, #_snhacungcap').select2({
                width: "100%",
                language: {
                    noResults: function () {
                        return "Không tìm thấy";
                    }
                }
            });
            $(document).on('click', '#btnThemCTNK', function () {
                if(!$("#btnSua").is(":disabled")){
                    alert('Vui lòng nhấn nút Sửa ở trên trước');
                    return false;
                }
                $('#modalThemNKCT').modal('show');
                $('#modalThemNKCT').find('input, select').val('');

                return false;
            });
            $(document).on('click','#btnLuuThemCTNK',function(){
                var obj = new Array();
                obj[0] = $('#_ctbienlai').val();
                obj[1] = $('#_ctsoluong').val();
                obj[2] = $('#_ctgianhap').val();
                obj[3] = $('#_ctthanhtien').val();
                obj[4] = document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value;
                for (var i = obj.length - 1; i >= 0; i--) {
                    if (obj[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                $.ajax({
                    url: 'dmnhapkhoNC.aspx/ThemNhapKhoCT',
                    type: 'POST',
                    data: "{obj:" + JSON.stringify(obj) + "}",
                    contentType: "application/json; charset=utf-8",
                    success: function (data) {
                        if ($.parseJSON(data.d)) {
                            alert('Thêm nhập kho chi tiết thành công');
                            $('#btnSua').attr('disabled', 'disabled');
                            Grid2.refresh();
                            $('#modalThemNKCT').modal('hide');
                        }
                        else {
                            alert('Thêm chi tiết nhập kho không thành công, kiểm tra lại thông tin');
                        }
                    }
                });
                return false;
            });
            $(document).on('click','#btnLuuSuaCTNK',function(){
                var obj = new Array();
                obj[0] = $('#_sctbienlai').val();
                obj[1] = $('#_sctsoluong').val();
                obj[2] = $('#_sctgianhap').val();
                obj[3] = $('#_sctthanhtien').val();
                obj[4] = $('#_sctsttNKCTpr').val();;
                for (var i = obj.length - 1; i >= 0; i--) {
                    if (obj[i].length==0) {
                        alert('Vui lòng điền đầy đủ thông tin');
                        return false;
                    }
                }
                $.ajax({
                    url: 'dmnhapkhoNC.aspx/SuaNhapKhoCT',
                    type: 'POST',
                    data: "{obj:" + JSON.stringify(obj) + "}",
                    contentType: "application/json; charset=utf-8",
                    success: function (data) {
                        if ($.parseJSON(data.d)) {
                            alert('Sửa nhập kho chi tiết thành công');
                            Grid2.refresh();
                            $('#modalSuaNKCT').modal('hide');
                        }
                        else {
                            alert('Sửa chi tiết nhập kho không thành công, kiểm tra lại thông tin');
                        }
                    }
                });
                return false;
            });
        });
        window.onload = function () {
            Grid1.refresh();
            Grid2.refresh();
        }
        function suaNhapKho(id) {
            if ($.isEmptyObject(id)) return false;
            $('#khung1').hide(234);
            $('#khung2').hide(234);
            $('#khung3').show(234);
            $('#btnSua').prop('disabled', false);
            var date;
            $.ajax({
                url: 'dmnhapkhoNC.aspx/LayThongTinSuaNhapKho',
                type: 'POST',
                data: "{obj:" + JSON.stringify(id) + "}",
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    var record = $.parseJSON(data.d)[0];
                    $('#_ssophieu').val(record.soPhieu);
                    $('#_ssophieu').attr('disabled', 'disabled');
                    $('#_snguoinhap').val(record.maNhanVienpr_sd).change();
                    $('#_snhacungcap').val(record.maNhaCCpr).change();
                    $('#_snoidung').val(record.noiDung);
                    $('#_sghichu').val(record.ghiChu);
                    document.getElementById('<%= txtHDLoaiG2.ClientID%>').value = "loadgridtuSttNK";
                    document.getElementById('<%= txtSttNKpr_sd.ClientID%>').value = record.sttNKpr;
                    Grid2.refresh();
                    try {
                        date = record.ngayNhap.split("-");
                        $('#_sngaynhap').val(date[0] + "-" + date[1] + "-" + date[2].substr(0, 2)).change();
                    } catch (e) {
                        document.getElementById('_sngaynhap').value = null;
                    }
                }
            });
            return false;
        }
        function suaNhapKhoCT(id) {
            if ($.isEmptyObject(id)) return false;
            if(!$("#btnSua").is(":disabled")){
                alert('Vui lòng nhấn nút Sửa ở trên trước');
                return false;
            }
            $.ajax({
                url: 'dmnhapkhoNC.aspx/LayThongTinSuaNhapKhoCT',
                type: 'POST',
                data: "{obj:" + JSON.stringify(id) + "}",
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    var record = $.parseJSON(data.d)[0];
                    $('#_sctbienlai').val(record.sttBienLaipr_sd).change();
                    $('#_sctsoluong').val(record.soLuong);
                    $('#_sctgianhap').val(record.donGia);
                    $('#_sctthanhtien').val(record.thanhTien);
                    $('#_sctsttNKCTpr').val(record.sttNKCTpr);
                    $('#modalSuaNKCT').modal('show');
                }
            });
            return false;
        }
        function xoaNhapKho(id){
            if ($.isEmptyObject(id)) return false;
            var yn = confirm('Bạn có thực sự muốn xóa nhập kho này?');
            if (!yn) return false;
            $.ajax({
                url: 'dmnhapkhoNC.aspx/XoaNhapKho',
                type: 'POST',
                data: "{obj:" + JSON.stringify(id) + "}",
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    if ($.parseJSON(data.d)) {
                        alert('Đã xóa nhập kho thành công');
                    }else{
                        alert('Không thể xóa, phiếu nhập kho này đang được sử dụng');
                    }
                }
            });
            Grid1.refresh();
            return false;
        }
        function xoaNhapKhoCT(id){
            if ($.isEmptyObject(id)) return false;
            var yn = confirm('Bạn có thực sự muốn xóa nhập kho chi tiết này?');
            if (!yn) return false;
            $.ajax({
                url: 'dmnhapkhoNC.aspx/XoaNhapKhoCT',
                type: 'POST',
                data: "{obj:" + JSON.stringify(id) + "}",
                contentType: "application/json; charset=utf-8",
                success: function (data) {
                    if ($.parseJSON(data.d)) {
                        alert('Đã xóa nhập kho chi tiết thành công');
                    }else{
                        alert('Không thể xóa phiếu nhập kho chi tiết này');
                    }
                }
            });
            Grid2.refresh();
            return false;
        }
        function TinhThanhTien() {
            try {
                var soluong = parseInt($('#_ctsoluong').val());
                var gianhap = parseInt($('#_ctgianhap').val());
                if (isNaN((soluong*gianhap).toString())) {
                    $('#_ctthanhtien').val("0");
                }else{
                    $('#_ctthanhtien').val((soluong*gianhap).toString());
                }
                
            } catch (e) {
                $('#_ctthanhtien').val("0");
            }
            return false;
        }
        function TinhThanhTienS() {
            try {
                var soluong = parseInt($('#_sctsoluong').val());
                var gianhap = parseInt($('#_sctgianhap').val());
                if (isNaN((soluong*gianhap).toString())) {
                    $('#_sctthanhtien').val("0");
                }else{
                    $('#_sctthanhtien').val((soluong*gianhap).toString());
                }
                
            } catch (e) {
                $('#_sctthanhtien').val("0");
            }
            return false;
        }
    </script>
        <script>
        var searchTimeout = null;
        function searchValue(grid, index, value) {
            if (searchTimeout != null) {
                return false;
            }
            if (jQuery.type(value) == "undefined")
                value = '';
            for (var i = index; i < grid.ColumnsCollection.length; i++) {
                if (grid.ColumnsCollection[i].HeaderText != "") {
                    var s = grid.ColumnsCollection[i].DataField;
                    if (grid.ColumnsCollection[i].Visible == true && s != "") {
                        grid.addFilterCriteria(s, OboutGridFilterCriteria.Contains, value);
                        console.log(s);
                    }
                }
            }
            searchTimeout = window.setTimeout(grid.executeFilter(), 2000);
            searchTimeout = null;
            return false;
        }
    </script>
</asp:Content>

