<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="~/Default.Master" CodeBehind="dmhuyentinh.aspx.cs" 
    Inherits="NTSoftReceipt.danhmuc.dmhuyentinh" %>
<%@ Register Assembly="obout_ComboBox" Namespace="Obout.ComboBox" TagPrefix="cc3" %>
<%@ Register Assembly="obout_Interface" Namespace="Obout.Interface" TagPrefix="cc2" %>
<%@ Register Assembly="obout_Window_NET" Namespace="OboutInc.Window" TagPrefix="owd" %>
<%@ Register Assembly="obout_Grid_NET" Namespace="Obout.Grid" TagPrefix="cc1" %>
<asp:Content ID="content1" ContentPlaceHolderID="head" runat="server">
    <script type="text/javascript">
        function ChonTinh() {
            document.getElementById('<%= _txtHiddenMatinh.ClientID %>').value=this.value();
            Grid1.refresh();
        }
        function LamTuoiLuoi() {
            Grid1.refresh();
        }
    </script>
</asp:Content>
<asp:Content ID="content2" ContentPlaceHolderID="combobox1" runat="server">
    <cc3:ComboBox runat="server" ID="Combobox1">
        <ClientSideEvents OnSelectedIndexChanged="ChonTinh" />
    </cc3:ComboBox>
    <!-- Trường ẩn -->
    <asp:HiddenField runat="server" ID="_txtHiddenMatinh" />
    
</asp:Content>

<asp:Content ID="content3" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <cc1:Grid runat="server" ID="Grid1" 
        FolderStyle="~/App_Themes/Styles/style_7" AllowPaging="true" 
        PageSizeOptions="5,10,15,20" PageSize="15" 
        AutoGenerateColumns="false"
        AllowFiltering="true" 
        EnableRecordHover="true" 
        AllowGrouping="false" 
        Height="350" OnRebind="Grid1_Rebind">
        <PagingSettings Position="Bottom" />
        <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
        <ClientSideEvents  />
        <Columns>
            <cc1:Column AllowDelete="true" AllowEdit="true" HeaderText="Sửa/Xóa" Width="100px">
            </cc1:Column>
            <cc1:Column HeaderText="Mã" DataField="maHuyenpr" Width="120px" ReadOnly="true" Wrap ="true"
                Visible="false">
            </cc1:Column>
            <cc1:Column HeaderText="Tên Quận/Huyện" DataField="tenHuyen" Width="300px" Wrap ="true">
            </cc1:Column>
            <cc1:Column HeaderText="Tên Tỉnh" DataField="tenTinh" Width="300px" Wrap ="true">
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
                    <b>Quận/Huyện - Tỉnh</b>
                </Template>
            </cc1:GridTemplate>
        </Templates>
    </cc1:Grid>
</asp:Content>