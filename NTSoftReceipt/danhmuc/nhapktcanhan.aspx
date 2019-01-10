<%@ Page Title="" Language="C#" MasterPageFile="~/ThiDuaKhenThuong/_frmThiDuaKhenThuong.Master"
    AutoEventWireup="true" CodeBehind="nhapktcanhan.aspx.cs" Inherits="QLNS2014.ThiDuaKhenThuong.nhapktcanhan" %>

<%@ Register Assembly="obout_Calendar2_Net" Namespace="OboutInc.Calendar2" TagPrefix="obout" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="ajaxToolkit" %>
<%@ Register Assembly="obout_Window_NET" Namespace="OboutInc.Window" TagPrefix="owd" %>
<%@ Register Assembly="obout_Interface" Namespace="Obout.Interface" TagPrefix="cc1" %>
<%@ Register Assembly="obout_ComboBox" Namespace="Obout.ComboBox" TagPrefix="cc2" %>
<%@ Register Assembly="obout_Grid_NET" Namespace="Obout.Grid" TagPrefix="cc3" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <style type="text/css">
        .style1 {
            font-size: medium;
            font-weight: bold;
        }

        .style2 {
            width: 100%;
        }

        div#_ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_Calendar3Container {
            z-index: 300000 !important;
        }

        div#_ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_Calendar5Container {
            z-index: 300000 !important;
        }

        div#_ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_Calendar6Container {
            z-index: 300000 !important;
        }
    </style>
    <style type="text/css">
        .item {
            position: relative !important;
            display: -moz-inline-stack;
            display: inline-block;
            zoom: 1;
            *display: inline;
            overflow: hidden;
            white-space: nowrap;
        }

        .header {
            margin-left: 2px;
        }

        .c1 {
            width: 320px;
        }

        .c2 {
            margin-left: 5px;
            width: 100px;
        }

        .c3 {
            display: none;
        }

        div#ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_cbdonViCongTac_ob_CbocbdonViCongTacItemsContainer {
            z-index: 2000000;
        }

        div#ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_cboDanhXung_ob_CbocboDanhXungItemsContainer {
            z-index: 2000000;
        }

        div#ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_cboPhongBan_DT_ob_CbocboPhongBan_DTItemsContainer {
            z-index: 2000000;
        }

        div#ctl00_ContentPlaceHolder1_windowThemMoiDoiTuong_cboGioiTinh_ob_CbocboGioiTinhItemsContainer {
            z-index: 2000000;
        }
    </style>

    <script type="text/javascript">
        var cot1;
        var cot2;
        var cot3;
        var coThaoTac = "";
        var daLuu = "0";
        var tendoituong = "";
        /////////////Xu ly moi
        window.onload = function () {
            $('#dinhKemQuyetDinh').html(' <a href="#" onclick="chonQuyetDinh(); return false;">Kèm quyết định</a>');
            loadCBoSearch();
            loadDoiTen();
        }
        function nhapDauKy() {
            //ngaydangky.value(QLNS2014.ThiDuaKhenThuong.nhapktcanhan.layNgayHT().value);
            //if (chkDauKy.checked() == true) {
            //    if (rdThiDua.checked() == true)
            //    {
            //    document.getElementById("lbngaydangky").style.visibility = "visible";
            //    document.getElementById("lbngaydangky").style.display = "Block";
            //}
            //}
            //else {
            //    document.getElementById("lbngaydangky").style.visibility = "hidden";
            //    document.getElementById("lbngaydangky").style.display = "none";
            //}
            Grid6.refresh();
            return false;
        }
        function loadCbo() {
            if (rdKhenThuong.checked()) {
                document.getElementById("danhHieu").innerHTML = "Hình thức khen thưởng";
                document.getElementById("loaiHinhKT").innerHTML = "Loại hình khen thưởng";
            }
            else {
                document.getElementById("danhHieu").innerHTML = "Danh hiệu thi đua";
                document.getElementById("loaiHinhKT").innerHTML = "Hình thức tổ chức thi đua";
            }
            //Load cbo danh hiệu
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.dsDanhHieuKTAll((rdKhenThuong.checked() ? "0" : "1"));
            cboDanhHieuWD1.options.clear();
            for (var i = 0; i < result.value.Rows.length; i++) {
                var ten = result.value.Rows[i].tenKhenThuong + "";
                var ma = result.value.Rows[i].maKhenThuongpr + "";
                cboDanhHieuWD1.options.add(ten, ma, i);
            }
            //Load cbo hình thức
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.dsHinhThucKTAll((rdKhenThuong.checked() ? "2" : "1"));
            cboHinhThucWD1.options.clear();
            for (var i = 0; i < result.value.Rows.length; i++) {
                var ten = result.value.Rows[i].tenHinhThucTDKT + "";
                var ma = result.value.Rows[i].maHinhThucTDKTpr + "";
                cboHinhThucWD1.options.add(ten, ma, i);
            }
            cboDanhHieuWD1.selectedIndex(0);
            cboHinhThucWD1.selectedIndex(0);

        }
        function loadCBoSearch() {
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.dsDanhHieuKT((rdKhenThuong.checked() ? "0" : "1"));
            cboSearchDanhHieu.options.clear();
            for (var i = 0; i < result.value.Rows.length; i++) {
                var ten = result.value.Rows[i].tenKhenThuong + "";
                var ma = result.value.Rows[i].maKhenThuongpr + "";
                cboSearchDanhHieu.options.add(ten, ma, i);
            }
            cboSearchDanhHieu.selectedIndex(0);
        }
        function chonKhenThuong() {
            loadCBoSearch();
            document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = "";
            Grid1.refresh();
            Grid4.refresh();
            loadDoiTen();
            return false;
        }
        function chonDonVi() {
            Grid1.refresh();
            return false;
        }
        function chonDonVi_WD() {
            Grid6.refresh();

            return false;
        }
        function chonDanhHieu() {
            Grid1.refresh();
            return false;
        }

        function moWindow(key) {
            document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTTDKT").value = key;
            Grid2.refresh();
            Dialog1.Open();
            return false;
        }
        function xoaFile(record) {
            QLNS2014.ThiDuaKhenThuong.nhapktcanhan.xoaFileDinhKem(record.sttHinhAnhpr, record.duongDan);
            Grid2.refresh();
            return false;
        }

        function chonQuyetDinh() {
            if (daLuu == "0") {
                alert("Bạn phải lưu thông tin quyết định thi đua khen thưởng trước");
                return false;
            }
            Window3.Open(); Window3.screenCenter(); return false;
        }

        function getDateNow() {
            var today = new Date();
            var dd = today.getDate();
            var mm = today.getMonth() + 1;
            var yyyy = today.getFullYear();
            if (dd < 10) {
                dd = '0' + dd
            }
            if (mm < 10) {
                mm = '0' + mm
            }
            return (dd + "/" + mm + "/" + yyyy);
        }
        function anHienControl(flag) {
            if (flag == "t") {
                $('#thongTinKTCaNhan').hide();
                $('#chiTietKTCaNhan').show();
                coThaoTac = "0";
                ngayKhenThuong.value(getDateNow());
                cboDanhHieuWD1.value('');
                cboHinhThucWD1.value('');
                moTa.value('');
                soQuyetDinh.value('');
                ngayKy.value('');
                cboCapQuyetDinh.value('');
                nguoiKy.value('');
                trichYeu.value('');
                noiDung.value('');
                cboHoiDong.value('');
                cboNamXetKhenThuong.selectedIndex(0);
                cboHoiDong.enable();
                ngayKhenThuong.enable();
                cboDanhHieuWD1.enable();
                cboHinhThucWD1.enable();
                moTa.enable();
                soQuyetDinh.enable();
                ngayKy.enable();
                cboCapQuyetDinh.enable();
                nguoiKy.enable();
                trichYeu.enable();
                noiDung.enable();
                btnLuuVaDong.enable();
                cboLinhVuc.value('');
                document.getElementById("ctl00_ContentPlaceHolder1_hdfFileUpload").value = "";
                daLuu = "0";
            }
            else if (flag == "s") {
                daLuu = "0";
                coThaoTac = "1";
                $('#thongTinKTCaNhan').hide();
                $('#chiTietKTCaNhan').show();
                cboDanhHieuWD1.enable();;
                cboHinhThucWD1.enable();
                soQuyetDinh.enable();
                ngayKy.enable();
                cboHoiDong.enable();
                cboCapQuyetDinh.enable();
                nguoiKy.enable();
                btnLuuVaDong.enable();
                trichYeu.enable();
                noiDung.enable();
                ngayKhenThuong.enable();
                moTa.enable();
                cboLinhVuc.enable();
            }
            else if (flag == "l") {
                coThaoTac = "";
                daLuu = "1";
                ngayKhenThuong.disable();
                cboDanhHieuWD1.disable();
                cboHinhThucWD1.disable();
                cboHoiDong.disable();
                moTa.disable();
                soQuyetDinh.disable();
                ngayKy.disable();
                cboCapQuyetDinh.disable();
                nguoiKy.disable();
                trichYeu.disable();
                noiDung.disable();
                btnLuuVaDong.disable();
                cboLinhVuc.disable();
            }
            else if (flag == "p") {
                loadCbo();
                $('#thongTinKTCaNhan').show();
                $('#chiTietKTCaNhan').hide();
                daLuu = "0";
                coThaoTac = "";
                btnLuuVaDong.enable();
                document.getElementById("ctl00_ContentPlaceHolder1_hdfFileUpload").value = "";
                loadCBoSearch();
                document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = "";
                Grid1.refresh();
                Grid4.refresh();
                return false;
            }
            return false;
        }
        function btnThemClick() {
            var permiss = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[6] == "true") {
                var permiss_ = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KiemTraKhoaChucNang();
                if (permiss_.value == true) {
                    alert("User đang sử dụng đã bị khóa không thực hiện thao tác được!");
                    return false;
                }
            }
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác thêm!");
                return false;
            }
            loadCbo();
            anHienControl("t");
            ctl00_ContentPlaceHolder1_Calendar4.enabled = true;
            ctl00_ContentPlaceHolder1_Calendar2.enabled = true;
            themMoiHoiDong.enable();
            cboLinhVuc.enable();
            document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = "";
            Grid3.refresh();
            $('#dinhKemQuyetDinh').html(' <a href="#" onclick="chonQuyetDinh(); return false;">Kèm quyết định</a>');
            loadComBoxHoiDong();
            return false;
        }
        function btnLuuVaDongClick() {
            // debugger;
            //            if (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgay(ngayKhenThuong.value()).value == false) {
            //                alert("Ngày khen thưởng phải là 10 ký tự dạng dd/MM/yyyy");
            //                return false;
            //            }
            if (cboNamXetKhenThuong.value() == "") {
                alert("Bạn chưa chọn năm xét khen thưởng");
                return false;
            }
            if (cboDanhHieuWD1.value() == "") {
                alert("Bạn chưa chọn danh hiệu thi đua");
                return false;
            }
            if (cboHinhThucWD1.value() == "") {
                alert("Bạn chưa chọn hình thức tổ chức thi đua");
                return false;
            }
            if (soQuyetDinh.value() == "") {
                alert("Bạn chưa nhập số quyết định khen thưởng");
                return false;
            }
            if (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgay(ngayKy.value()).value == false) {
                alert("Ngày ký quyết định phải là 10 ký tự dạng dd/MM/yyyy");
                return false;
            }
            if (cboCapQuyetDinh.value() == "") {
                alert("Bạn chưa chọn cấp quyết định khen thưởng");
                return false;
            }
            if (cboHoiDong.value() == "") {
                alert("Bạn chưa chọn hội đồng xét duyệt");
                return false;
            }
            if (cboLinhVuc.value() == "") {
                alert("Bạn chưa chọn lĩnh vực");
                return false;
            }

            var param = new Array();
            param[0] = document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value;
            param[1] = moTa.value();
            param[2] = soQuyetDinh.value();
            param[3] = ngayKy.value();
            param[4] = cboCapQuyetDinh.value();
            param[5] = noiDung.value();
            param[6] = cboDanhHieuWD1.value();
            param[7] = trichYeu.value();
            param[8] = nguoiKy.value();
            param[9] = cboHinhThucWD1.value();
            param[10] = ngayKhenThuong.value();
            param[11] = cboHoiDong.value();
            param[12] = cboNamXetKhenThuong.value();
            param[13] = cboLinhVuc.value();
            // debugger;
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.luuThongTin(param, coThaoTac);
            if (result.value == "0") {
                alert("Thông tin quyết định khen thưởng chưa được lưu");
                return false;
            }
            if (coThaoTac == "0")
                document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = result.value;
            daLuu = "1";
            anHienControl("l");
            ctl00_ContentPlaceHolder1_Calendar4.enabled = false;
            ctl00_ContentPlaceHolder1_Calendar2.enabled = false;
            themMoiHoiDong.disable();

            return false;
        }
        function Grid3OnBeforeClientAdd() {
            if (daLuu == "0" || daLuu == "") {
                alert("Bạn phải lưu thông tin quyết định khen thưởng trước!");
                return false;
            }
            Grid6.refresh();
            setWindowPosition();
            Window2.Open();
            return false;
        }
        function btnChonDoiTuongClick() {
            var doiTuong = "";
            var dangKy = "";
            if (document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value == "0") {
                alert("Thông tin chung chưa được lưu, vui lòng thao tác lại");
                return false;
            }
            //if (chkDauKy.checked() == true) {
            //    var ktngayKT = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgay(ngaydangky.value());
            //    if (ktngayKT.value == false) {
            //        alert("Ngày ký phải có dạng dd/MM/yyyy");
            //        ngaydangky.focus();
            //        return false;
            //    }
            //}
            //var ktngayKT = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgay(ngaydangky.value());
            //if (ktngayKT.value == false) {
            //    alert("Ngày ký phải có dạng dd/MM/yyyy");
            //    ngaydangky.focus();
            //    return false;
            //}
            if (Grid6.SelectedRecords.length > 0) {
                for (var i = 0; i < Grid6.SelectedRecords.length; i++) {
                    var record = Grid6.SelectedRecords[i];
                    if (chkDauKy.checked() == true) {
                        var kiemtra = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KiemTraDoiTuongTDKT(record.sttDoiTuongTDKTpr_sd, cboDanhHieuWD1.value()).value;
                        if (kiemtra != "") {
                            alert("Cá nhân đã được khen thưởng theo số quyết định " + kiemtra);
                            return false;
                        }
                    }
                    if (record.ketQua == "Không đủ điều kiện") {
                        alert("Cá nhân [" + record.tenDoiTuongTDKT + "] không đủ điều kiện xét duyệt danh hiệu [" + cboDanhHieuWD1.options[cboDanhHieuWD1.selectedIndex()].text + "]");
                        return false;
                    }
                    doiTuong += "'" + record.sttDoiTuongTDKTpr_sd + "',"
                    dangKy += "'" + record.sttTDKTpr + "',"
                }
            } else {
                alert("Không có dòng dữ liệu nào được chọn!");
            }
            var param = new Array();
            param[0] = moTa.value();
            param[1] = soQuyetDinh.value();
            param[2] = ngayKy.value();
            param[3] = cboCapQuyetDinh.value();
            param[4] = noiDung.value();
            param[5] = cboDanhHieuWD1.value();
            param[7] = trichYeu.value();
            param[6] = document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value;
            param[8] = nguoiKy.value();
            param[9] = cboHinhThucWD1.value();
            param[10] = ngayKhenThuong.value();
            param[11] = dangKy.substring(0, dangKy.length - 1);
            param[12] = doiTuong.substring(0, doiTuong.length - 1);
            param[13] = chkDauKy.checked();
            param[14] = cboNamXetKhenThuong.value();
            //param[14] = ngaydangky.value();
            // param[15] = cboHoiDong.value();
            param[15] = cboLinhVuc.value();
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.chonDoiTuong(param, coThaoTac);
            if (result.value == "0") {
                alert("Thông tin nhập kết quả TĐKT chưa được lưu");
                return false;
            }

            Window2.Close();
            Grid3.refresh();
            for (i = 0; i < Grid6.Rows.length; i++) {
                Grid6.deselectRecord(i);
            }
            return false;
        }
        function Grid1_OnClientSelect(record) {
            loadCbo();
            document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = record[0].sttQuyetDinhKTpr;
            Grid4.refresh();
            return false;
        }
        function Grid1_OnClientDblClick(record) {
            var permiss = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[6] == "true") {
                var permiss_ = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KiemTraKhoaChucNang();
                if (permiss_.value == true) {
                    alert("User đang sử dụng đã bị khóa không thực hiện thao tác được!");
                    return false;
                }
            }
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác sửa!");
                return false;
            }
            coThaoTac = "1";
            moTa.value('');
            soQuyetDinh.value('');
            ngayKy.value('');
            cboCapQuyetDinh.value('');
            noiDung.value('');
            cboDanhHieuWD1.value('');
            trichYeu.value('');
            cboHoiDong.value('');
            nguoiKy.value('');
            cboHinhThucWD1.value('');
            ngayKhenThuong.value('');

            document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = Grid1.Rows[record].Cells["sttQuyetDinhKTpr"].Value;
            moTa.value(Grid1.Rows[record].Cells["moTa"].Value);
            soQuyetDinh.value(Grid1.Rows[record].Cells["soQuyetDinh"].Value);
            ngayKy.value(Grid1.Rows[record].Cells["ngayKy"].Value);
            cboCapQuyetDinh.value(Grid1.Rows[record].Cells["capQuyetDinh"].Value);
            noiDung.value(Grid1.Rows[record].Cells["noiDung"].Value);
            cboDanhHieuWD1.value(Grid1.Rows[record].Cells["maKhenThuongpr_sd"].Value);
            trichYeu.value(Grid1.Rows[record].Cells["trichYeu"].Value);

            nguoiKy.value(Grid1.Rows[record].Cells["nguoiKy"].Value);
            cboHinhThucWD1.value(Grid1.Rows[record].Cells["maHinhThucTDKTpr_sd"].Value);
            ngayKhenThuong.value(Grid1.Rows[record].Cells["ngayKhenThuong"].Value);
            cboNamXetKhenThuong.value(Grid1.Rows[record].Cells["namXetKhenThuong"].Value);
            loadComBoxHoiDong();
            cboHoiDong.value(Grid1.Rows[record].Cells["sttHoiDongXDTDpr_sd"].Value);
            cboLinhVuc.value(Grid1.Rows[record].Cells["maLinhVucpr_sd"].Value);
            if (Grid1.Rows[record].Cells["fileDinhKem"].Value != "") {
                $('#dinhKemQuyetDinh').html('<a href="#" onclick="taiQuyetDinh(\'' + Grid1.Rows[record].Cells["fileDinhKem"].Value + '\')"; return false;">Tải quyết định khen thưởng</a>   <a href="#" onclick="xoaQuyetDinhKT(' + Grid1.Rows[record].Cells["sttQuyetDinhKTpr"].Value + ',\'' + Grid1.Rows[record].Cells["fileDinhKem"].Value + '\'); return false;">Xóa </a>');
            }
            else {
                $('#dinhKemQuyetDinh').html(' <a href="#" onclick="chonQuyetDinh(); return false;">Kèm quyết định</a>');
            }
            Grid3.refresh();
            anHienControl("s");
            daLuu = "1";

            return false;
        }
        function taiQuyetDinh(fileName) {
            window.open(QLNS2014.ThiDuaKhenThuong.nhapktcanhan.taiFile(fileName).value);
            return false;
        }
        function xoaQuyetDinhKT(sttQuyetDinhKTpr, filename) {
            if (daLuu == "0") {
                alert("Bạn phải lưu thông tin quyết định khen thưởng trước!");
                return false;
            }
            QLNS2014.ThiDuaKhenThuong.nhapktcanhan.xoaQuyetDinh(sttQuyetDinhKTpr + "", filename);
            $('#dinhKemQuyetDinh').html(' <a href="#" onclick="chonQuyetDinh(); return false;">Kèm quyết định</a>');
            return false;
        }
        function btnSuaClick() {
            var permiss = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[6] == "true") {
                var permiss_ = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KiemTraKhoaChucNang();
                if (permiss_.value == true) {
                    alert("User đang sử dụng đã bị khóa không thực hiện thao tác được!");
                    return false;
                }
            }
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác sửa!");
                return false;
            }
            loadCbo();
            if (Grid1.SelectedRecords.length > 0) {
                for (var i = 0; i < Grid1.SelectedRecords.length; i++) {
                    var record = Grid1.SelectedRecords[i];
                    document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = record.sttQuyetDinhKTpr;

                    loadCbo();
                    moTa.value(record.moTa);
                    soQuyetDinh.value(record.soQuyetDinh);
                    ngayKy.value(record.ngayKy);
                    cboCapQuyetDinh.value(record.capQuyetDinh);

                    noiDung.value(record.noiDung);
                    cboDanhHieuWD1.value(record.maKhenThuongpr_sd);
                    trichYeu.value(record.trichYeu);
                    nguoiKy.value(record.nguoiKy);
                    cboHinhThucWD1.value(record.maHinhThucTDKTpr_sd);
                    ngayKhenThuong.value(record.ngayKhenThuong);
                    cboNamXetKhenThuong.value(record.namXetKhenThuong);
                    loadComBoxHoiDong();
                    cboHoiDong.value(record.sttHoiDongXDTDpr_sd);
                    if (record.fileDinhKem != "") {
                        $('#dinhKemQuyetDinh').html('<a href="#" onclick="taiQuyetDinh(\'' + record.fileDinhKem + '\')"; return false;">Tải quyết định khen thưởng</a>   <a href="#" onclick="xoaQuyetDinhKT(' + record.sttQuyetDinhKTpr + ',\'' + record.fileDinhKem + '\'); return false;">Xóa </a>');
                    }
                    else {
                        $('#dinhKemQuyetDinh').html(' <a href="#" onclick="chonQuyetDinh(); return false;">Kèm quyết định</a>');
                    }
                    cboLinhVuc.value(record.maLinhVucpr_sd);
                    anHienControl("s");
                    coThaoTac = "1";
                    Grid3.refresh();
                    daLuu = "1";

                    return false;
                }

            } else {
                alert("Không có dòng dữ liệu nào được chọn!");
            }
        }
        function beforeDelete(record) {
            document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTTDKT").value = record.sttTDKTpr;
            idDialogYesNo.Open();
            //           QLNS2014.ThiDuaKhenThuong.nhapktcanhan.xoaDoiTuongTDKT(record.sttTDKTpr);
            //           Grid3.refresh();beforeUpdate
            return false;
        }
        function beforeUpdate(record) {
            //  debugger;
            var param = new Array();
            param[0] = record.ngayDangKy;
            param[1] = record.sttTDKTpr;
            param[2] = cboPhongBan.value();
            param[3] = txtThamNienQD.value();
            param[4] = txtGhiChu.value();
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.capNhatNgayDK(param);
            if (!result) {
                alert("Thông tin cập nhật ngày đăng ký TĐKT chưa được lưu");
                return false;
            }
            Grid3.refresh();
            return false;
        }
        function reBindAlphabe(control, text) {
            var arr = new Array();
            arr = 'btnLinkTatCa,LinkButton1,LinkButton4,LinkButton5,LinkButton6,LinkButton8,LinkButton9,LinkButton10,LinkButton11,LinkButton12,LinkButton13,LinkButton14,LinkButton15,LinkButton16,LinkButton17,LinkButton18,LinkButton19,LinkButton20,LinkButton21,LinkButton22,LinkButton23,LinkButton24,LinkButton25,LinkButton26,LinkButton27,LinkButton28,LinkButton29'.split(',');
            for (var i = 0; i < arr.length; i++) {
                var control1 = arr[i].toString();
                if (arr[i].toString() == control)
                    document.getElementById("ctl00_ContentPlaceHolder1_" + control).style.backgroundColor = "lightblue";
                else
                    document.getElementById("ctl00_ContentPlaceHolder1_" + control1).style.backgroundColor = "";
            }
            document.getElementById("ctl00_ContentPlaceHolder1_hdfAlphabet").value = text;
            //alert(document.getElementById('hdfAlphabet').value);
            Grid1.refresh();
            return false;
        }
        function reBindAlphabe1(control, text) {
            var arr = new Array();
            arr = 'LinkButton54,LinkButton55,LinkButton56,LinkButton30,LinkButton31,LinkButton32,LinkButton33,LinkButton34,LinkButton35,LinkButton36,LinkButton37,LinkButton38,LinkButton39,LinkButton40,LinkButton41,LinkButton42,LinkButton43,LinkButton44,LinkButton45,LinkButton46,LinkButton47,LinkButton48,LinkButton49,LinkButton50,LinkButton51,LinkButton52,LinkButton53'.split(',');
            document.getElementById("ctl00_ContentPlaceHolder1_hdfAlphabet1").value = text;
            Grid6.refresh();
            return false;
        }
        function truocKhiXoa() {
            var permiss = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[6] == "true") {
                var permiss_ = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KiemTraKhoaChucNang();
                if (permiss_.value == true) {
                    alert("User đang sử dụng đã bị khóa không thực hiện thao tác được!");
                    return false;
                }
            }
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác xóa!");
                return false;
            }
            if (document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value != "" || document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value != "0") {
                var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraXoa("sttQuyetDinhKTpr_sd", document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value).value
                if (result == "") {
                    Dialog2.Open();
                }
                else {
                    $('.msgDialog').html("Quyết định khen thưởng đã được sử dụng trong <br/>" + result);
                    msgDialog.Open();
                }
                return false;
            }
            else {
                alert('Không tồn tại dữ liệu muốn xóa');
                return false;
            }
        }
        function xoaTTQuyetDinhKT() {
            var ktraXoa = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.xoaQuyetDinhKT(document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value);
            if (ktraXoa.value == false) {
                alert("Xóa thông tin thất bại");
                return false;
            }
            document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value = "0";
            Grid1.refresh();
            Grid4.refresh();
            Dialog2.Close();
        }
        function xoaDuLieu() {
            var ktraXoa = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.xoaDoiTuongTDKT(document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTTDKT").value);
            if (ktraXoa.value == false) {
                alert("Xóa thông tin thất bại");
                return false;
            }
            Grid3.refresh();
            idDialogYesNo.Close();
        }
        function showPopup() {
            $('#hinhAnhSlide').show();
            $('[data-popup="popup-1"]').fadeIn(350);
            $(function () {
                //----- OPEN
                $('[data-popup-open]').on('click', function (e) {
                    var targeted_popup_class = jQuery(this).attr('data-popup-open');
                    $('[data-popup="' + targeted_popup_class + '"]').fadeIn(350);
                    e.preventDefault();
                });
                //----- CLOSE
                $('[data-popup-close]').on('click', function (e) {
                    var targeted_popup_class = jQuery(this).attr('data-popup-close');
                    $('[data-popup="' + targeted_popup_class + '"]').fadeOut(350);
                    e.preventDefault();
                });
            });
            return false;
        }
        function setWindowPosition() {
            var screenWidth = screen.width;
            var screenHeight = screen.height;
            var Window2Size = Window2.getSize();
            Window2.setPosition(parseFloat((parseFloat(screenWidth) - parseFloat(Window2Size.width)) / 2), 50);
        }

        function ThayDoiGiaTriSo(sender, numDec) {
            if (((numDec < 48 || numDec > 57) && (numDec < 96 || numDec > 105)) && numDec != 191 && numDec != 8 && numDec != 37 && numDec != 39 && numDec != 13 && numDec != 111 && numDec != 9 && numDec != 17 && numDec != 67 && numDec != 86 && numDec != 88) {
                return false;
            }
        }
        function uploadExcelError(sender, args) {

            alert("Tải tập tin Excel mẫu thất bại!");
        }

        function uploadExcelComplete(sender, args) {
            var kq = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KTraFile());
            var fileExtension = args.get_fileName();

            if (fileExtension.indexOf('.xls') != -1) {

                if (kq == false) {
                    alert("Tập tin Excel mẫu không đúng cấu trúc");
                    return false;
                } else {
                    if (chkNhapDauKy.checked() == true) {
                        var tontai = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.soLuongTonTai());
                        if (tontai.value == "khongco") {
                            var kq = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.NhapDuLieuFileExcel());
                            document.getElementById("info").innerHTML = kq.value;
                            for (i = 0; i < Grid1.Rows.length; i++) {
                                Grid1.deselectRecord(i);
                            }
                            Grid1.SelectedRecordsContainer.value = "";
                            Grid1.SelectedRecords = new Array();
                            Grid1.refresh();
                        }
                        else {
                            DialogYesNoNhapExcel.Open();
                            document.getElementById("noidungkiemtraexcel").innerHTML = tontai.value;
                            return false;
                        }
                    }
                    else {
                        var kq = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.NhapDuLieuFileExcel_DangKy());
                        document.getElementById("info").innerHTML = kq.value;
                        for (i = 0; i < Grid1.Rows.length; i++) {
                            Grid1.deselectRecord(i);
                        }
                        Grid1.SelectedRecordsContainer.value = "";
                        Grid1.SelectedRecords = new Array();
                        Grid1.refresh();
                    }
                }

            } else {
                document.getElementById("info").innerHTML = "<br/>Tập tin Excel mẫu không đúng định dạng.";
                return;
            }
            document.getElementById('processExcel').style.visibility = "hidden";
        }
        function hienThongBaoLoi() {
            Dialog_ThongBaoLoi.screenCenter();
            Dialog_ThongBaoLoi.Open();
            var danhsachdoituong_ = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.laydanhsachdoituong());
            document.getElementById("noidungloi").innerHTML = danhsachdoituong_.value;
            return false;
        }
        function uploadExcelStarted() {
            document.getElementById("info").innerHTML = "<br/>Đang tải tệp tin...";
            document.getElementById('processExcel').style.visibility = "visible";
        }
        function loadDoiTen() {
            if (rdKhenThuong.checked()) {
                Grid1.GridHeaderContainer.firstChild.firstChild.firstChild.firstChild.childNodes[3].firstChild.firstChild.innerHTML = "Hình thức khen thưởng";
                Grid1.GridHeaderContainer.firstChild.firstChild.firstChild.firstChild.childNodes[4].firstChild.firstChild.innerHTML = "Loại hình khen thưởng";
            }
            else {
                Grid1.GridHeaderContainer.firstChild.firstChild.firstChild.firstChild.childNodes[3].firstChild.firstChild.innerHTML = "Danh hiệu thi đua";
                Grid1.GridHeaderContainer.firstChild.firstChild.firstChild.firstChild.childNodes[4].firstChild.firstChild.innerHTML = "Hình thức tổ chức thi đua";
            }
        }
        function truockhisuachitiet(record) {
            //dtNgayDangKy.value(record.ngayDangKy);
            loadComBoxPhongBan(record.maDonVipr_noicongtac);
            cboPhongBan.value(record.sttDoiTuongTDKTpr_cha);
        }
        function loadComBoxHoiDong() {
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.loadComBoxHoiDong(cboNamXetKhenThuong.value());
            cboHoiDong.options.clear();
            for (var i = 0; i < result.value.Rows.length; i++) {
                var ten = result.value.Rows[i].tenHoiDong;
                var ma = result.value.Rows[i].sttHoiDongXDTDpr;
                cboHoiDong.options.add(ten, ma, i);
            }
            // loadcboHDSKKN();
        }
        function loadcboHDSKKN() {
            //                var a = ngayKhenThuong.value();
            //                var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.loadcboHDSKKN(a);
            //                cboHoiDong.options.clear();
            //                for (var i = 0; i < result.value.Rows.length; i++) {
            //                    var sttHDSKKN = result.value.Rows[i].sttHoiDongXDSKKNpr + "";
            //                    var tenHoiDong = result.value.Rows[i].tenHoiDong;
            //                    cboHoiDong.options.add(tenHoiDong, sttHDSKKN, i);
            //                }

        }
        function themMoiHoiDong() {
            cbthuocDonVi.value(cbdonViCongTacGrid.value());
            taoMaDoiTuong1();
            coThaotac = 1;
            _txtDoiTuongTT.value('');
            _txtNgayThanhLap.value('');
            _txtKyHieu.value('');
            _cboDoiTuongTapThe.value('');
            checkNgoaiDV.checked(false);
            Window4.screenCenter();
            Window4.Open();
            return false;
        }
        function showhoidong() {
            txtTenHoiDong.value('');
            txtNgayLap.value('');
            txtGhiChu.value('');
            var namThaoTac = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.getNamThaoTac();
            cboNamXetDuyet.value(namThaoTac.value);
            Window4.screenCenter();
            Window4.Open();
            return false;
        }
        function isEmptyText(value) {
            if (value == null) {
                return true;
            }
            else {
                var len = value.replace(new RegExp(" ", "g"), '').length;
                if (len > 0)
                    return false;
                else
                    return true;
            }
        }
        function LuuvaDongSKKN() {
            var param = new Array();
            param[0] = txtTenHoiDong.value();
            param[1] = txtNgayLap.value();
            param[2] = cboNamXetDuyet.value();
            param[3] = txtGhiChu.value();
            // param[4] = document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTHoiDongXD").value;
            if (isEmptyText(param[0])) {
                alert("Tên hội đồng không được bỏ trống!");
                setTimeout(function () { txtTenHoiDong.focus(); }, 100);
                return false;
            }
            if (isEmptyText(param[1])) {
                alert("Ngày lập không được bỏ trống!");
                setTimeout(function () { txtNgayLap.focus(); }, 100);
                return false;
            }
            if (isEmptyText(param[1])) {
                alert("Ngày lập không được bỏ trống!");
                setTimeout(function () { txtNgayLap.focus(); }, 100);
                return false;
            }
            var ktNgay = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgay(param[1]);
            if (ktNgay.value == false) {
                alert("Ngày lập phải có 10 ký tự dạng dd/MM/yyyy");
                setTimeout(function () { txtNgayLap.focus(); }, 100);
                return false;
            }
            var stt = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.insertGrid1(param).value;
            daLuu = "1";
            if (stt != '0')
                cboHoiDong.options.add(txtTenHoiDong.value(), stt, cboHoiDong.length);
            else {
                alert("Thêm hội đồng Thi đua khen thưởng không thành công.");
            }
            Window4.Close();
            return false;
        }
        function loadComBoxPhongBan(maDonVipr_sd) {
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.loadComBoxPhongBan(maDonVipr_sd).value;
            cboPhongBan.options.clear();
            for (var i = 0; i < result.Rows.length; i++) {
                var ten = result.Rows[i].tenDoiTuongTDKT;
                var ma = result.Rows[i].sttDoiTuongTDKTpr;
                cboPhongBan.options.add(ten, ma, i);
            }
        }
        //function loadComBoxPhongBan(maDonVipr_sd) {
        //    var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.loadComBoxPhongBan(maDonVipr_sd).value;
        //    cboPhongBan.options.clear();
        //    for (var i = 0; i < result.value.Rows.length; i++) {
        //        var ten = result.Rows[i].tenDoiTuongTDKT;
        //        var ma = result.Rows[i].sttDoiTuongTDKTpr;
        //        cboPhongBan.options.add(ten, ma, i);
        //    }
        //}
        function truocKhiThemMoiDoiTuong() {
            _txtTenDoiTuong.value('');
            txtNgayVeCQ.value('');
            txtChucVuHienTai.value('');
            cboDanhXung.value('0');
            cboGioiTinh.value('');
            _txtNgaySinh.value('');
            _txtCMND.value('');
            _txtNgayCap.value('');
            _txtNoiCap.value('');
            cbdonViCongTac.value(cboDonVi_WD.value());
            cbdonViCongTac.disable();
            cboPhongBan_DT.selectedIndex = 0;
            chkLaLanhDao.checked(false);
            txtMaDoiTuong.value(QLNS2014.ThiDuaKhenThuong.nhapktcanhan.taoMaDoiTuong1(cboDonVi_WD.value()).value);
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.loadComBoxPhongBan(cboDonVi_WD.value()).value;

            cboPhongBan_DT.options.clear();
            for (var i = 0; i < result.Rows.length; i++) {
                var ten = result.Rows[i].tenDoiTuongTDKT;
                var ma = result.Rows[i].sttDoiTuongTDKTpr;
                cboPhongBan_DT.options.add(ten, ma, i);
            }
            windowThemMoiDoiTuong.Open();
            windowThemMoiDoiTuong.screenCenter();
        }
        function windowThemMoiDoiTuong_open() {
            //debugger;
            var result = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.loadComBoxPhongBan(cbdonViCongTac.value());
            cboPhongBan_DT.options.clear();
            //for (var i = 0; i < result.value.Rows.length; i++) {
            //    var ten = result.Rows[i].tenDoiTuongTDKT;
            //    var ma = result.Rows[i].sttDoiTuongTDKTpr;
            //    cboPhongBan_DT.options.add(ten, ma, i);
            //}
            return false;
        }
        function layGioiTinh() {
            // debugger;
            if (cboDanhXung.value() == 'Ông')
                cboGioiTinh.value('Nam');
            if (cboDanhXung.value() == 'Bà')
                cboGioiTinh.value('Nữ');
            return false;
        }
        function LuuvaDongDoiTuong() {
            var pra = new Array();
            pra[0] = _txtTenDoiTuong.value();
            pra[1] = txtChucVuHienTai.value();
            pra[2] = cboGioiTinh.value();
            pra[3] = _txtNgaySinh.value();
            pra[4] = _txtCMND.value();
            pra[5] = _txtNgayCap.value();
            pra[6] = _txtNoiCap.value();
            pra[7] = cboPhongBan_DT.value();
            pra[8] = cboDanhXung.value();
            pra[9] = txtMaDoiTuong.value();
            pra[10] = chkLaLanhDao.checked();
            pra[11] = cbdonViCongTac.value();
            pra[12] = txtNgayVeCQ.value();
            var ngayCap = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.ktNgayBC(pra[5]);
            if (isEmptyText(pra[9]) == true) {
                alert("Mã cá nhân không được trống");
                setTimeout(function () { txtMaDoiTuong.focus(); }, 100);
                return false;
            }
            if (isEmptyText(pra[0]) == true) {
                alert("Tên cá nhân không được trống");
                setTimeout(function () { _txtTenDoiTuong.focus(); }, 100);
                return false;
            }
            if (isEmptyText(pra[11]) == true) {
                alert("Đơn vị công tác không được trống");
                setTimeout(function () { cbdonViCongTac.focus(); }, 100);
                return false;
            }
            if (txtNgayVeCQ.value() != "") {
                if (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgay(txtNgayVeCQ.value()).value == false) {
                    alert("Ngày về cơ quan phải là 10 ký tự định dạng dd/MM/yyyy");
                    return false;
                }
            }
            if (_txtNgaySinh.value() != '') {

                if (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraNgayNam(_txtNgaySinh.value()).value == false) {
                    alert("Ngày sinh phải là năm hoặc 10 ký tự định dạng dd/MM/yyyy");
                    setTimeout(function () { _txtNgaySinh.focus(); }, 100);
                    return false;
                }
            }
            if (isEmptyText(pra[5]) == false) {
                if (ngayCap.value == false) {
                    alert("Ngày cấp phải là 10 ký tự định dạng dd/MM/yyyy");
                    return false;
                }
            }
            var ktMaDoiTuong = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.ktraMaDoiTuong(pra[9], pra[11]);
            if (ktMaDoiTuong.value != "0") {
                alert("Mã cá nhân đã tồn tại");
                return false;
            }
            // kiểm tra họ tên
            var ktHoTenNamSinh = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.ktraHoTenVaNamSinh(pra[0], pra[3], pra[11]);
            if (ktHoTenNamSinh.value == "0") {
                alert("Tên cán bộ này đã tồn tại");
                return false;
            }
            if (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.LuuCaNhan(pra).value == false)
                alert("Thêm cá nhân thi đua, khen thưởng không thành công.")
            Grid6.refresh();
            windowThemMoiDoiTuong.Close();
            return false;
        }

        function btnXuatWord_Click() {
            if (Grid1.SelectedRecords.length > 0) {
                var record = Grid1.SelectedRecords[0];
                var kq = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.XuatWord(record.sttQuyetDinhKTpr).value;
                if (kq != "") {
                    window.open(kq);
                }
            } else {
                alert("Không có dòng dữ liệu nào được chọn");
            }
            return false;
        }
    </script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <asp:HiddenField ID="hdfAlphabet" runat="server" />
    <asp:HiddenField ID="hdfAlphabet1" runat="server" />
    <asp:HiddenField ID="hdfUrlFileWin1" runat="server" />
    <asp:HiddenField ID="hdfUrlFileWin2" runat="server" />
    <asp:HiddenField ID="hdfSTTTDKT" runat="server" />
    <asp:HiddenField ID="hdfFileUpload" runat="server" />
    <asp:HiddenField ID="sttQuyetDinhKTpr_sd" runat="server" Value="0" />
    <table style="width: 1000px">
        <tr>
            <td style="width: 50px">
                <asp:Image ID="Image1" runat="server" Height="40px" ImageUrl="~/Image/giayKhen.jpg"
                    Width="50px" />
            </td>
            <td style="font-size: 15px; font-weight: bold; color: #FE8900">NHẬP KHEN THƯỞNG CÁ NHÂN
            </td>
        </tr>
    </table>
    <div id="thongTinKTCaNhan">
        <div style="">
            <cc1:OboutButton ID="btnThem" runat="server" Text="Thêm" Width="80px" OnClientClick="btnThemClick();return false;"
                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton">
            </cc1:OboutButton>
            <cc1:OboutButton ID="btnSua" runat="server" Text="Sửa" Width="80px" OnClientClick="btnSuaClick();return false;">
            </cc1:OboutButton>
            <cc1:OboutButton ID="OboutButton3" runat="server" Text="Xóa" Width="80px" OnClientClick="truocKhiXoa(); return false;">
            </cc1:OboutButton>
            <cc1:OboutButton ID="btnNhapExcel" runat="server" Text="Nhập excel" Width="120px"
                OnClientClick="btnNhapExcel_Click();return false;" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton">
            </cc1:OboutButton>
            <cc1:OboutButton ID="btnTaiExcekMau" runat="server" Text="Tải excel mẫu" Width="120px"
                OnClientClick="btnTaiExcekMau_Click();return false;" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton">
            </cc1:OboutButton>
            <cc1:OboutButton ID="OboutButton2" runat="server" Text="Nhập word" Width="120px"
                OnClientClick="btnNhapWord_Click();return false;" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton">
            </cc1:OboutButton>
            <cc1:OboutButton ID="btnXuatWord" runat="server" Text="Xuất word" Width="120px"
                OnClientClick="btnXuatWord_Click();return false;" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton">
            </cc1:OboutButton>
        </div>
        <div style="margin-bottom: 2px; margin-top: 4px">
            <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
                <legend><b>Chức năng tìm kiếm</b></legend>
                <table style="width: 100%">
                    <tr>
                        <td>
                            <cc1:OboutRadioButton ID="rdKhenThuong" GroupName="a" runat="server" Text="Khen thưởng"
                                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutRadioButton">
                                <ClientSideEvents OnCheckedChanged="chonKhenThuong" />
                            </cc1:OboutRadioButton>
                            <cc1:OboutRadioButton ID="rdThiDua" GroupName="a" runat="server" Text="Thi đua" Checked="true"
                                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutRadioButton">
                            </cc1:OboutRadioButton>
                        </td>
                        <td style="width: 50px">Đơn vị
                        </td>
                        <td style="width: 250px" align="right">
                            <cc2:ComboBox ID="cboDonVi" runat="server" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox"
                                FilterType="Contains" Height="150" Width="100%" MenuWidth="150%">
                                <ClientSideEvents OnSelectedIndexChanged="chonDanhHieu" />
                            </cc2:ComboBox>
                        </td>
                        <td style="width: 180px">Danh hiệu thi đua, hình thức khen thưởng
                        </td>
                        <td style="width: 250px" align="right">
                            <cc2:ComboBox ID="cboSearchDanhHieu" runat="server" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox"
                                FilterType="Contains" Height="150" Width="100%" MenuWidth="150%">
                                <ClientSideEvents OnSelectedIndexChanged="chonDanhHieu" />
                            </cc2:ComboBox>
                        </td>
                    </tr>
                </table>
            </fieldset>
            <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
                <legend><b>Danh sách quyết định khen thưởng</b></legend>
                <cc3:Grid ID="Grid1" runat="server" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/style_7"
                    AllowPaging="true" PageSizeOptions="15,50,100,200,300,-1" PageSize="15" AutoGenerateColumns="false"
                    AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Height="300"
                    OnRebind="Grid1_OnRebind" Width="100%" AllowAddingRecords="false" AllowMultiRecordSelection="False">
                    <ScrollingSettings EnableVirtualScrolling="false" ScrollHeight="300" ScrollWidth="986" />
                    <PagingSettings Position="Bottom" />
                    <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
                    <ClientSideEvents OnClientSelect="Grid1_OnClientSelect" OnClientDblClick="Grid1_OnClientDblClick" />
                    <Columns>
                        <cc3:Column HeaderText="Ngày khen thưởng" DataField="ngayKhenThuong" DataFormatString="{0:dd/MM/yyyy}"
                            Width="150px" Wrap="true" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="Số quyết định" DataField="soQuyetDinh" Width="120px" Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Ngày quyết định" DataField="ngayKy" DataFormatString="{0:dd/MM/yyyy}"
                            Width="150px">
                        </cc3:Column>
                        <cc3:Column HeaderText="Danh hiệu thi đua, hình thức khen thưởng" DataField="tenKhenThuong"
                            Width="280px" Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Loại hình khen thưởng" DataField="hinhThucKhenThuong" Width="250px"
                            Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Trích yếu" DataField="trichYeu" Width="230px" Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Quyết định" DataField="fileDinhKem1" Width="120px" Wrap="true"
                            ParseHTML="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Năm xét KT" DataField="namXetKhenThuong" Width="100px" Visible="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Lĩnh vực" DataField="tenLinhVucKT" Width="200px" Visible="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="" DataField="fileDinhKem" Width="0px" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="moTa" DataField="moTa" Width="0px" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="noiDung" DataField="noiDung" Width="0px" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="capQuyetDinh" DataField="capQuyetDinh" Width="0px" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="nguoiKy" DataField="nguoiKy" Width="0px" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="sttQuyetDinhKTpr" DataField="sttQuyetDinhKTpr" Width="0px"
                            Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="maKhenThuongpr_sd" DataField="maKhenThuongpr_sd" Width="0px"
                            Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="maHinhThucTDKTpr_sd" DataField="maHinhThucTDKTpr_sd" Width="0px"
                            Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="sttHoiDongXDTDpr_sd" DataField="sttHoiDongXDTDpr_sd" Width="0px"
                            Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="maLinhVucpr_sd" DataField="maLinhVucpr_sd" Width="0px"
                            Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="" DataField="" Width="20px">
                        </cc3:Column>
                    </Columns>
                    <LocalizationSettings AddLink="Thêm mới" CancelAllLink="Hủy tất cả" CancelLink="Hủy"
                        DeleteLink="Xóa" Filter_ApplyLink="Tìm kiếm" Filter_HideLink="Đóng tìm kiếm"
                        EditLink="Học xong" Filter_RemoveLink="Xóa tìm kiếm" Filter_ShowLink="Mở tìm kiếm"
                        FilterCriteria_NoFilter="Không tìm kiếm" FilterCriteria_Contains="Chứa" FilterCriteria_DoesNotContain="Không chứa"
                        FilterCriteria_StartsWith="Bắt đầu với" FilterCriteria_EndsWith="Kết thúc với"
                        FilterCriteria_EqualTo="Bằng" FilterCriteria_NotEqualTo="Không bằng" FilterCriteria_SmallerThan="Nhỏ hơn"
                        FilterCriteria_GreaterThan="Lớn hơn" FilterCriteria_SmallerThanOrEqualTo="Nhỏ hơn hoặc bằng"
                        FilterCriteria_GreaterThanOrEqualTo="Lớn hơn hoặc bằng" FilterCriteria_IsNull="Rỗng"
                        FilterCriteria_IsNotNull="Không rỗng" FilterCriteria_IsEmpty="Trống" FilterCriteria_IsNotEmpty="Không trống"
                        Paging_OfText="của" Grouping_GroupingAreaText="" JSWarning="Có một lỗi khởi tạo lưới với ID '[GRID_ID]'. \ N \ n [Chú ý] \ n \ nHãy liên hệ bộ phận bảo trì của Nhất Tâm Soft để được giúp đỡ."
                        LoadingText="Đang tải...." MaxLengthValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX vượt quá số lượng tối đa ký tự YYYYY cho phép cột này."
                        ModifyLink="Chỉnh sửa" NoRecordsText="Không có dữ liệu" Paging_ManualPagingLink="Trang kế »"
                        Paging_PageSizeText="Số dòng 1 trang:" Paging_PagesText="Trang:" Paging_RecordsText="Dòng:"
                        ResizingTooltipWidth="Rộng:" SaveAllLink="Lưu tất cả" SaveLink="Lưu" TypeValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX là không đúng."
                        UndeleteLink="Không xóa" UpdateLink="Lưu" />
                </cc3:Grid>
            </fieldset>
            <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
                <legend><b>Chi tiết cá nhân được khen thưởng</b></legend>
                <cc3:Grid ID="Grid4" runat="server" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/style_7"
                    AllowPaging="true" PageSizeOptions="15,50,100,200,300,500,1000,-1" PageSize="15"
                    AutoGenerateColumns="false" AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false"
                    Height="350" OnRebind="Grid4_OnRebind" Width="100%" AllowAddingRecords="false"
                    AllowMultiRecordSelection="False">
                    <ScrollingSettings EnableVirtualScrolling="false" ScrollHeight="350" ScrollWidth="986" />
                    <PagingSettings Position="Bottom" />
                    <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
                    <Columns>
                        <cc3:Column HeaderText="STT" DataField="stt" Width="50px" Align="center">
                        </cc3:Column>
                        <cc3:Column HeaderText="Họ và tên" DataField="doiTuong" Width="240px" Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Chức vụ" DataField="chucVu" Width="150px" Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Ngày sinh" DataField="ngaySinh" Width="100px">
                        </cc3:Column>
                        <cc3:Column HeaderText="Ký hiệu" DataField="kyHieu" Width="0px" Visible="false">
                        </cc3:Column>
                        <cc3:Column HeaderText="Đơn vị công tác" DataField="noiCongTac" Width="266px" Wrap="true">
                        </cc3:Column>
                        <cc3:Column HeaderText="Hình ảnh" DataField="xemHinh" Width="80px" ParseHTML="true"
                            Wrap="true">
                        </cc3:Column>
                        <cc3:CheckBoxColumn DataField="checkDauKy" HeaderText="Khen thưởng năm trước" Width="120px"
                            Wrap="true">
                        </cc3:CheckBoxColumn>
                        <cc3:Column HeaderText="" DataField="" Width="20px">
                        </cc3:Column>
                    </Columns>
                    <LocalizationSettings AddLink="Thêm mới" CancelAllLink="Hủy tất cả" CancelLink="Hủy"
                        DeleteLink="Xóa" Filter_ApplyLink="Tìm kiếm" Filter_HideLink="Đóng tìm kiếm"
                        EditLink="Học xong" Filter_RemoveLink="Xóa tìm kiếm" Filter_ShowLink="Mở tìm kiếm"
                        FilterCriteria_NoFilter="Không tìm kiếm" FilterCriteria_Contains="Chứa" FilterCriteria_DoesNotContain="Không chứa"
                        FilterCriteria_StartsWith="Bắt đầu với" FilterCriteria_EndsWith="Kết thúc với"
                        FilterCriteria_EqualTo="Bằng" FilterCriteria_NotEqualTo="Không bằng" FilterCriteria_SmallerThan="Nhỏ hơn"
                        FilterCriteria_GreaterThan="Lớn hơn" FilterCriteria_SmallerThanOrEqualTo="Nhỏ hơn hoặc bằng"
                        FilterCriteria_GreaterThanOrEqualTo="Lớn hơn hoặc bằng" FilterCriteria_IsNull="Rỗng"
                        FilterCriteria_IsNotNull="Không rỗng" FilterCriteria_IsEmpty="Trống" FilterCriteria_IsNotEmpty="Không trống"
                        Paging_OfText="của" Grouping_GroupingAreaText="" JSWarning="Có một lỗi khởi tạo lưới với ID '[GRID_ID]'. \ N \ n [Chú ý] \ n \ nHãy liên hệ bộ phận bảo trì của Nhất Tâm Soft để được giúp đỡ."
                        LoadingText="Đang tải...." MaxLengthValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX vượt quá số lượng tối đa ký tự YYYYY cho phép cột này."
                        ModifyLink="Chỉnh sửa" NoRecordsText="Không có dữ liệu" Paging_ManualPagingLink="Trang kế »"
                        Paging_PageSizeText="Số dòng 1 trang:" Paging_PagesText="Trang:" Paging_RecordsText="Dòng:"
                        ResizingTooltipWidth="Rộng:" SaveAllLink="Lưu tất cả" SaveLink="Lưu" TypeValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX là không đúng."
                        UndeleteLink="Không xóa" UpdateLink="Lưu" />
                </cc3:Grid>
            </fieldset>
        </div>
    </div>
    <div id="chiTietKTCaNhan" style="display: none;">
        <table style="padding-bottom: 20px;">
            <tr>
                <td>
                    <cc1:OboutButton ID="btnLuuVaDong" runat="server" Text="Lưu" Width="100px" OnClientClick="btnLuuVaDongClick();return false;">
                    </cc1:OboutButton>
                    <cc1:OboutButton ID="btnDongWindow1" runat="server" Text="Quay ra" Width="100px"
                        OnClientClick="anHienControl('p'); return false;">
                    </cc1:OboutButton>
                </td>
            </tr>
        </table>
        <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
            <legend><b>Thông tin chung</b></legend>
            <table border="0" width="100%">
                <tr>
                    <td style="width: 160px; display: none">Ngày khen thưởng <span style="color: Red">[*]</span>
                    </td>
                    <td style="width: 310px; display: none">
                        <cc1:OboutTextBox ID="ngayKhenThuong" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"
                            runat="server" Width="275px">
                            <ClientSideEvents OnTextChanged="loadComBoxHoiDong" OnKeyDown="loadComBoxHoiDong" OnKeyPress="loadComBoxHoiDong" OnKeyUp="loadComBoxHoiDong" />
                        </cc1:OboutTextBox>
                        <obout:Calendar ID="Calendar4" runat="server" Columns="1" CultureName="vi-VN" DateFormat="dd/MM/yyyy"
                            DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                            DatePickerImageTooltip="Chọn ngày" DatePickerMode="true" MonthSelectorType="HtmlList"
                            TextBoxId="ngayKhenThuong" TimeSelectorType="HtmlList" YearSelectorType="HtmlList">
                        </obout:Calendar>
                    </td>
                    <td style="width: 160px">Năm xét khen thưởng <span style="color: Red">[*]</span>
                    </td>
                    <td style="width: 310px; text-align: right;">
                        <cc2:ComboBox ID="cboNamXetKhenThuong" runat="server" AppendDataBoundItems="false"
                            FilterType="Contains" Height="150" Width="100%" MenuWidth="100%">
                            <ClientSideEvents OnSelectedIndexChanged="loadComBoxHoiDong" OnTextChanged="loadComBoxHoiDong" />
                        </cc2:ComboBox>
                    </td>
                    <td style="padding-left: 12px">Lĩnh vực <span style="color: Red">[*]</span>
                    </td>
                    <td style="width: 310px; text-align: right;">
                        <cc2:ComboBox ID="cboLinhVuc" runat="server" FilterType="Contains" Height="150" Width="100%"
                            MenuWidth="100%">
                        </cc2:ComboBox>
                    </td>
                </tr>
                <tr>
                    <td>
                        <span id="danhHieu"></span><span style="color: Red">[*]</span>
                    </td>
                    <td>
                        <cc2:ComboBox ID="cboDanhHieuWD1" runat="server" AppendDataBoundItems="false" FilterType="Contains"
                            Height="150" Width="100%" MenuWidth="150%">
                        </cc2:ComboBox>
                    </td>
                    <td style="padding-left: 12px;">
                        <span id="loaiHinhKT"></span><span style="color: Red">[*]</span>
                    </td>
                    <td>
                        <cc2:ComboBox ID="cboHinhThucWD1" runat="server" AppendDataBoundItems="false" FilterType="Contains"
                            Height="150" Width="100%" MenuWidth="150%">
                        </cc2:ComboBox>
                    </td>
                </tr>
                <tr>
                    <td>Mô tả danh hiệu, thành tích đạt được
                    </td>
                    <td colspan="3">
                        <cc1:OboutTextBox ID="moTa" TextMode="MultiLine" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"></cc1:OboutTextBox>
                    </td>
                </tr>
                <tr>
                    <td>Số quyết định <span style="color: Red">[*]</span>
                    </td>
                    <td>
                        <cc1:OboutTextBox ID="soQuyetDinh" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"></cc1:OboutTextBox>
                    </td>
                    <td style="padding-left: 12px;">Ngày quyết định <span style="color: Red">[*]</span>
                    </td>
                    <td>
                        <cc1:OboutTextBox ID="ngayKy" runat="server" Width="275px">
                                <ClientSideEvents OnTextChanged="function(sender, value){thayDoiNgay(sender, value, 'LLS')}" OnKeyDown="ThayDoiGiaTriSo" />
                        </cc1:OboutTextBox>
                        <obout:Calendar ID="Calendar2" runat="server" Columns="1" CultureName="vi-VN" DateFormat="dd/MM/yyyy"
                            DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                            DatePickerImageTooltip="Chọn ngày" DatePickerMode="true" MonthSelectorType="HtmlList"
                            TextBoxId="ngayKy" TimeSelectorType="HtmlList" YearSelectorType="HtmlList">
                        </obout:Calendar>
                    </td>
                </tr>
                <tr>
                    <td>Cấp quyết định <span style="color: Red">[*]</span>
                    </td>
                    <td colspan="1">
                        <cc2:ComboBox ID="cboCapQuyetDinh" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox">
                        </cc2:ComboBox>
                    </td>
                    <td style="padding-left: 12px;">Người ký
                    </td>
                    <td>
                        <cc1:OboutTextBox ID="nguoiKy" runat="server" Width="100%"></cc1:OboutTextBox>
                    </td>
                </tr>
                <tr>
                    <td>Hội đồng <span style="color: Red">[*]</span>
                    </td>
                    <td colspan="3">
                        <cc2:ComboBox ID="cboHoiDong" runat="server" Height="100px" Width="93%" MenuWidth="100%"
                            FilterType="Contains" SelectedIndex="0">
                        </cc2:ComboBox>
                        <cc1:OboutButton ID="themMoiHoiDong" runat="server" Text="+" Width="53px" OnClientClick="showhoidong(); return false;"
                            UseSubmitBehavior="false">
                        </cc1:OboutButton>
                    </td>
                </tr>
                <tr>
                    <td>Trích yếu
                    </td>
                    <td colspan="3">
                        <cc1:OboutTextBox ID="trichYeu" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"></cc1:OboutTextBox>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top">Nội dung
                    </td>
                    <td colspan="3">
                        <cc1:OboutTextBox ID="noiDung" TextMode="MultiLine" runat="server" Width="100%">
                        </cc1:OboutTextBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="4">
                        <p id="dinhKemQuyetDinh">
                        </p>
                    </td>
                </tr>
            </table>
        </fieldset>
        <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
            <legend><b>Chi tiết cá nhân được khen thưởng</b></legend>
            <cc3:Grid ID="Grid3" runat="server" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/style_7"
                AllowPaging="true" PageSizeOptions="15,50,100,200,-1" PageSize="10" AutoGenerateColumns="false"
                AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" Height="300"
                OnRebind="Grid3_OnRebind" Width="976" ShowHeader="true" AllowAddingRecords="true"
                AllowMultiRecordSelection="false">
                <AddEditDeleteSettings AddLinksPosition="Top" />
                <ClientSideEvents OnBeforeClientDelete="beforeDelete" OnBeforeClientUpdate="beforeUpdate"
                    OnClientEdit="truockhisuachitiet" OnBeforeClientAdd="Grid3OnBeforeClientAdd" />
                <ScrollingSettings EnableVirtualScrolling="false" ScrollWidth="976" ScrollHeight="300" />
                <Columns>
                    <cc3:Column HeaderText="Thao tác" DataField="" AllowDelete="true" AllowEdit="true"
                        Width="80px">
                    </cc3:Column>
                    <cc3:Column HeaderText="Tên đối tượng" DataField="tenDoiTuongTDKT" Width="240px"
                        ReadOnly="true" Wrap="true">
                    </cc3:Column>
                    <cc3:Column HeaderText="Ngày đăng ký" DataField="ngayDangKy" Width="150px" Wrap="true"
                        ReadOnly="true" DataFormatString="{0:dd/MM/yyyy}">
                    </cc3:Column>
                    <cc3:Column HeaderText="Bộ phận công tác" DataField="phongBan" Width="180px" Wrap="true"
                        ApplyFormatInEditMode="true">
                        <TemplateSettings EditTemplateId="editPhongBan" />
                    </cc3:Column>
                    <cc3:Column HeaderText="Đơn vị công tác" DataField="donViCongTac" Width="280px" Wrap="true"
                        ReadOnly="true">
                    </cc3:Column>
                    <cc3:Column HeaderText="Hình ảnh" DataField="xemHinh" Width="100px" Wrap="true" ParseHTML="true"
                        ReadOnly="true">
                    </cc3:Column>
                    <cc3:CheckBoxColumn DataField="checkDauKy" HeaderText="Khen thưởng năm trước" Width="160px"
                        ReadOnly="true">
                    </cc3:CheckBoxColumn>
                    <cc3:Column HeaderText="Thâm niên quyết định" DataField="thamNienQD" Width="180px" Wrap="true"
                        ApplyFormatInEditMode="true">
                        <TemplateSettings EditTemplateId="editThamNienQD" />
                    </cc3:Column>
                    <cc3:Column HeaderText="Ghi chú" DataField="ghiChu" Width="250px" Wrap="true"
                        ApplyFormatInEditMode="true">
                        <TemplateSettings EditTemplateId="editGhiChu" />
                    </cc3:Column>
                    <cc3:Column HeaderText="sttTDKTpr" DataField="sttTDKTpr" Visible="false" Width="0px">
                    </cc3:Column>
                    <cc3:Column HeaderText="sttDoiTuongTDKTpr_cha" DataField="sttDoiTuongTDKTpr_cha"
                        Visible="false" Width="0px">
                    </cc3:Column>
                    <cc3:Column HeaderText="maDonVipr_noicongtac" DataField="maDonVipr_noicongtac" Visible="false"
                        Width="0px">
                    </cc3:Column>
                    <cc3:Column HeaderText="" DataField="" Width="20px">
                    </cc3:Column>
                    <cc3:Column HeaderText="maLinhVucpr_sd" DataField="maLinhVucpr_sd" Visible="false"
                        Width="0px">
                    </cc3:Column>
                </Columns>
                <LocalizationSettings AddLink="Thêm mới" CancelAllLink="Hủy tất cả" CancelLink="Hủy"
                    DeleteLink="Xóa" Filter_ApplyLink="Tìm kiếm" Filter_HideLink="Đóng tìm kiếm"
                    EditLink="Sửa" Filter_RemoveLink="Xóa tìm kiếm" Filter_ShowLink="Mở tìm kiếm"
                    FilterCriteria_Contains="Contain" Grouping_GroupingAreaText="" JSWarning="Có một lỗi khởi tạo lưới với ID '[GRID_ID]'. \ N \ n [Chú ý] \ n \ nHãy liên hệ bộ phận bảo trì của Nhất Tâm Soft để được giúp đỡ."
                    LoadingText="Đang tải...." MaxLengthValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX vượt quá số lượng tối đa ký tự YYYYY cho phép cột này."
                    ModifyLink="Chỉnh sửa" NoRecordsText="Không có dữ liệu" Paging_ManualPagingLink="Trang kế »"
                    Paging_PageSizeText="Số dòng 1 trang:" Paging_PagesText="Trang:" Paging_RecordsText="Dòng:"
                    ResizingTooltipWidth="Rộng:" SaveAllLink="Lưu tất cả" SaveLink="Lưu" TypeValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX là không đúng."
                    UndeleteLink="Không xóa" UpdateLink="Lưu" />
                <Templates>
                    <cc3:GridTemplate runat="server" ID="editPhongBan" ControlID="cboPhongBan" ControlPropertyName="value">
                        <Template>
                            <cc2:ComboBox runat="server" ID="cboPhongBan" Width="100%" MenuWidth="350" Height="150"
                                FilterType="Contains" FolderStyle="App_Themes/Styles/Interface/OboutComboBox"
                                AppendDataBoundItems="false">
                            </cc2:ComboBox>
                        </Template>
                    </cc3:GridTemplate>
                    <cc3:GridTemplate runat="server" ID="tmpngayDangKy" ControlID="dtNgayDangKy" ControlPropertyName="value">
                        <Template>
                            <cc1:OboutTextBox ID="dtNgayDangKy" runat="server" Width="73%">
                                <ClientSideEvents OnTextChanged="function(sender, value){thayDoiNgay(sender, value, 'LLS')}" OnKeyDown="ThayDoiGiaTriSo" />
                            </cc1:OboutTextBox>
                            <obout:Calendar ID="Calendar2" runat="server" Columns="1" CultureName="vi-VN" DateFormat="dd/MM/yyyy"
                                DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                                OnClientDateChanged="CalendarW3_Changed" DatePickerImageTooltip="Chọn ngày" DatePickerMode="true"
                                MonthSelectorType="HtmlList" TextBoxId="dtNgayDangKy" TimeSelectorType="HtmlList"
                                YearSelectorType="HtmlList">
                            </obout:Calendar>
                        </Template>
                    </cc3:GridTemplate>
                    <cc3:GridTemplate runat="server" ID="editThamNienQD" ControlID="txtThamNienQD" ControlPropertyName="value">
                        <Template>
                            <cc1:OboutTextBox ID="txtThamNienQD" runat="server" Width="100%"></cc1:OboutTextBox>
                        </Template>
                    </cc3:GridTemplate>
                    <cc3:GridTemplate runat="server" ID="editGhiChu" ControlID="txtGhiChu" ControlPropertyName="value">
                        <Template>
                            <cc1:OboutTextBox ID="txtGhiChu" runat="server" Width="100%"></cc1:OboutTextBox>
                        </Template>
                    </cc3:GridTemplate>
                </Templates>
            </cc3:Grid>
            <asp:SqlDataSource ID="dtsPhongBan" runat="server"></asp:SqlDataSource>

            <script type="text/javascript">
                function CalendarW3_Changed(sender, selectedDate) {
                    dtNgayDangKy.value(ChuyenNgay(selectedDate.toString()));
                }
                function ChuyenNgay(date) {
                    var reval = "";
                    if (date != null) {
                        var newDate = new Date(date);
                        var dd = newDate.getDate();
                        var mm = newDate.getMonth() + 1;
                        var yyyy = newDate.getFullYear();
                        var ldd, lmm;
                        ldd = dd.toString().length;
                        lmm = mm.toString().length;
                        if (ldd <= 1)
                            dd = '0' + dd.toString();
                        if (lmm <= 1)
                            mm = '0' + (mm);
                        reval = dd + '/' + mm + '/' + yyyy;
                    }
                    return reval;
                }
            </script>

        </fieldset>

        <script>
            //Xử lý mới
            function uploadPDFError(sender, args) {
                document.getElementById("msg").innerHTML = "Tải tập tin thất bại!";
                document.getElementById('process').style.visibility = "hidden";
                document.getElementById("ctl00_ContentPlaceHolder1_Window3_AsyncFileUpload1_ctl02").value = "";
            }
            function uploadPDFComplete(sender, args) {
                var fileExtension = args.get_fileName();
                if (fileExtension.indexOf('.pdf') != -1) {
                    document.getElementById("msg").innerHTML = "";
                    Window3.Close();
                    document.getElementById("ctl00_ContentPlaceHolder1_Window3_AsyncFileUpload1_ctl02").value = "";
                    $('#dinhKemQuyetDinh').html('<a href="#" onclick="taiQuyetDinh(\'' + QLNS2014.ThiDuaKhenThuong.nhapktcanhan.getPDFFileName().value + '\')"; return false;">Tải quyết định khen thưởng</a>   <a href="#" onclick="xoaQuyetDinhKT(' + document.getElementById("ctl00_ContentPlaceHolder1_sttQuyetDinhKTpr_sd").value + ',\'' + QLNS2014.ThiDuaKhenThuong.nhapktcanhan.getPDFFileName().value + '\'); return false;">   Xóa </a>');
                    document.getElementById('process').style.visibility = "hidden";
                }
                else {
                    document.getElementById("msg").innerHTML = "Tệp tin không đúng định dạng. Chỉ được tải file có dạng .pdf!";
                    document.getElementById("ctl00_ContentPlaceHolder1_Window3_AsyncFileUpload1_ctl02").value = "";
                    document.getElementById('process').style.visibility = "hidden";
                    return;
                }
            }
            function uploadPDFStarted() {
                document.getElementById("msg").innerHTML = "<br/>Đang tải tệp tin...";
                document.getElementById('process').style.visibility = "visible";
            }
        </script>

        <owd:Window ID="Window3" runat="server" IsModal="true" ShowCloseButton="true" Title="Chọn quyết định khen thưởng"
            RelativeElementID="WindowPositionHelper" Height="150" Width="400" VisibleOnLoad="false"
            StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma" IsResizable="false"
            ShowStatusBar="false">
            <table>
                <tr>
                    <td style="width: 453px" colspan="3">
                        <ajaxToolkit:AsyncFileUpload OnClientUploadError="uploadPDFError" OnClientUploadComplete="uploadPDFComplete"
                            OnClientUploadStarted="uploadPDFStarted" runat="server" ID="AsyncFileUpload1"
                            UploaderStyle="Traditional" UploadingBackColor="" ThrobberID="myThrobber" BorderStyle="NotSet"
                            Font-Underline="False" Font-Strikeout="False" />
                        <span><i>Chỉ cho phép tải lên file có định dạng: pdf</i></span><br />
                        <span id="msg" style="color: Red"></span>
                        <div id="process" style="visibility: hidden; display: block;">
                            <img alt="" src="/ThiDuaKhenThuong/images/029.gif" width="120px" />
                        </div>
                    </td>
                </tr>
            </table>
        </owd:Window>
    </div>
    <owd:Window ID="Window2" runat="server" IsModal="true" ShowCloseButton="true" ShowStatusBar="false"
        Title="Chọn cá nhân" RelativeElementID="WindowPositionHelper" Height="580" Width="800"
        VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        IsResizable="false">
        <table border="0px" style="margin-top: 3px; margin-bottom: 3px; width: 100%;">
            <tr>
                <td style="width: 170px">
                    <cc1:OboutButton ID="btnChonDoiTuong" runat="server" Text="Chọn" Width="80px" OnClientClick="btnChonDoiTuongClick(); return false;">
                    </cc1:OboutButton>
                    <cc1:OboutButton ID="OboutButton1" runat="server" Text="Đóng" Width="80px" OnClientClick="Window2.Close(); return false;">
                    </cc1:OboutButton>
                    <cc1:OboutButton ID="OboutButton4" runat="server" Text="Thêm mới đối tượng" Width="135px"
                        OnClientClick="truocKhiThemMoiDoiTuong(); return false;">
                    </cc1:OboutButton>
                </td>
                <td style="width: 250px; text-align: right;">
                    <cc1:OboutCheckBox ID="chkDauKy" Checked="false" runat="server" Text="Nhập khen thưởng các năm trước"
                        FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutCheckBox">
                        <ClientSideEvents OnCheckedChanged="nhapDauKy" />
                    </cc1:OboutCheckBox>
                </td>
                <%--<td id="lbngaydangky"  style="width: 350px; text-align: right;display:none;visibility:hidden">
                     <div>
                        Ngày ký <span style="color: Red">[*]</span> <cc1:OboutTextBox ID="ngaydangky" runat="server" Width="220px" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"></cc1:OboutTextBox>
                            <obout:Calendar ID="Calendar3" runat="server" Columns="1" CultureName="vi-VN" DateFormat="dd/MM/yyyy"
                                DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                                DatePickerImageTooltip="Chọn ngày" DatePickerMode="true" MonthSelectorType="HtmlList"
                                TextBoxId="ngaydangky" TimeSelectorType="HtmlList" YearSelectorType="HtmlList">
                            </obout:Calendar></div>
                    </td>--%>
            </tr>
        </table>
        <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
            <legend><b>Chức năng tìm kiếm</b></legend>
            <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                    <td style="width: 80px">Đơn vị
                    </td>
                    <td>
                        <cc2:ComboBox ID="cboDonVi_WD" runat="server" Height="150" FilterType="Contains"
                            Width="100%" AutoClose="true" MenuWidth="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox">
                            <ClientSideEvents OnSelectedIndexChanged="chonDonVi_WD" />
                        </cc2:ComboBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <asp:LinkButton ID="LinkButton3" runat="server" OnClientClick="reBindAlphabe1('LinkButton54','');return false;"><b>Tất cả</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton57" runat="server" OnClientClick="reBindAlphabe1('LinkButton55','A');return false;"
                            Width="7px"><b>A</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton58" runat="server" OnClientClick="reBindAlphabe1('LinkButton56','B');return false;"
                            Width="7px"><b>B</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton59" runat="server" OnClientClick="reBindAlphabe1('LinkButton30','C');return false;"
                            Width="7px"><b>C</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton60" runat="server" OnClientClick="reBindAlphabe1('LinkButton31','D');return false;"
                            Width="7px"><b>D</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton61" runat="server" OnClientClick="reBindAlphabe1('LinkButton32','E');return false;"
                            Width="7px"><b>E</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton62" runat="server" OnClientClick="reBindAlphabe1('LinkButton33','F');return false;"
                            Width="7px"><b>F</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton63" runat="server" OnClientClick="reBindAlphabe1('LinkButton34','G');return false;"
                            Width="7px"><b>G</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton64" runat="server" OnClientClick="reBindAlphabe1('LinkButton35','H');return false;"
                            Width="7px"><b>H</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton65" runat="server" OnClientClick="reBindAlphabe1('LinkButton36','I');return false;"
                            Width="7px"><b>I</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton66" runat="server" OnClientClick="reBindAlphabe1('LinkButton37','J');return false;"
                            Width="7px"><b>J</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton67" runat="server" OnClientClick="reBindAlphabe1('LinkButton38','K');return false;"
                            Width="7px"><b>K</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton68" runat="server" OnClientClick="reBindAlphabe1('LinkButton39','L');return false;"
                            Width="7px"><b>L</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton69" runat="server" OnClientClick="reBindAlphabe1('LinkButton40','M');return false;"
                            Width="7px"><b>M</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton70" runat="server" OnClientClick="reBindAlphabe1('LinkButton41','N');return false;"
                            Width="7px"><b>N</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton71" runat="server" OnClientClick="reBindAlphabe1('LinkButton42','O');return false;"
                            Width="7px"><b>O</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton72" runat="server" OnClientClick="reBindAlphabe1('LinkButton43','P');return false;"
                            Width="7px"><b>P</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton73" runat="server" OnClientClick="reBindAlphabe1('LinkButton44','Q');return false;"
                            Width="7px"><b>Q</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton74" runat="server" OnClientClick="reBindAlphabe1('LinkButton45','R');return false;"
                            Width="7px"><b>R</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton75" runat="server" OnClientClick="reBindAlphabe1('LinkButton46','S');return false;"
                            Width="7px"><b>S</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton76" runat="server" OnClientClick="reBindAlphabe1('LinkButton47','T');return false;"
                            Width="7px"><b>T</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton77" runat="server" OnClientClick="reBindAlphabe1('LinkButton48','U');return false;"
                            Width="7px"><b>U</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton78" runat="server" OnClientClick="reBindAlphabe1('LinkButton49','V');return false;"
                            Width="7px"><b>V</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton79" runat="server" OnClientClick="reBindAlphabe1('LinkButton50','W');return false;"
                            Width="7px"><b>W</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton80" runat="server" OnClientClick="reBindAlphabe1('LinkButton51','X');return false;"
                            Width="7px"><b>X</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton81" runat="server" OnClientClick="reBindAlphabe1('LinkButton52','Y');return false;"
                            Width="7px"><b>Y</b></asp:LinkButton>
                        <b>|</b>
                        <asp:LinkButton ID="LinkButton82" runat="server" OnClientClick="reBindAlphabe1('LinkButton53','Z');return false;"
                            Width="7px"><b>Z</b></asp:LinkButton>
                    </td>
                </tr>
            </table>
        </fieldset>
        <br />
        <fieldset style="border: 1px solid #DBDBE1; margin: 3px; padding: 4px">
            <legend><b>Danh sách cá nhân</b></legend>
            <cc3:Grid ID="Grid6" runat="server" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/style_7"
                AllowPaging="true" PageSizeOptions="100,300,500,-1" PageSize="300" AutoGenerateColumns="false"
                AllowFiltering="true" EnableRecordHover="true" AllowGrouping="false" OnRebind="Grid6_OnRebind"
                Height="385" Width="790px" ShowHeader="true" AllowAddingRecords="false" AllowMultiRecordSelection="false">
                <ScrollingSettings EnableVirtualScrolling="false" ScrollHeight="430" ScrollWidth="770" />
                <ClientSideEvents OnClientDblClick="Grid2_OnClientDblClick" />
                <PagingSettings Position="Bottom" />
                <FilteringSettings InitialState="Hidden" FilterPosition="Top" FilterLinksPosition="Bottom" />
                <Columns>
                    <cc3:CheckBoxSelectColumn Width="50px">
                    </cc3:CheckBoxSelectColumn>
                    <cc3:Column HeaderText="sttTDKTpr" DataField="sttTDKTpr" Visible="false">
                    </cc3:Column>
                    <cc3:Column HeaderText="sttDoiTuongTDKTpr_sd" DataField="sttDoiTuongTDKTpr_sd" Visible="false">
                    </cc3:Column>
                    <cc3:Column HeaderText="Mã cá nhân" DataField="maDoiTuong" Width="120px" Visible="false">
                    </cc3:Column>
                    <cc3:Column HeaderText="Cá nhân" DataField="tenDoiTuongTDKT" Width="220px">
                    </cc3:Column>
                    <cc3:Column HeaderText="Ngày sinh" DataField="ngaySinh" DataFormatString="{0:dd/MM/yyyy}"
                        Width="120px" Visible="true">
                    </cc3:Column>
                    <cc3:Column HeaderText="Đơn vị công tác" DataField="donViCongTac" Width="250px" Visible="true">
                    </cc3:Column>
                    <cc3:Column HeaderText="Danh hiệu đăng ký thi đua" DataField="tenKhenThuong" Width="250px"
                        Visible="true">
                    </cc3:Column>
                    <cc3:Column HeaderText="Kết quả" DataField="ketQua" Width="250px" Visible="true">
                    </cc3:Column>
                    <cc3:Column HeaderText="" DataField="" Width="20px" Visible="false">
                    </cc3:Column>
                </Columns>
                <LocalizationSettings AddLink="Thêm mới" CancelAllLink="Hủy tất cả" CancelLink="Hủy"
                    DeleteLink="Xóa" Filter_ApplyLink="Tìm kiếm" Filter_HideLink="Đóng tìm kiếm"
                    EditLink="Học xong" Filter_RemoveLink="Xóa tìm kiếm" Filter_ShowLink="Mở tìm kiếm"
                    FilterCriteria_NoFilter="Không tìm kiếm" FilterCriteria_Contains="Chứa" FilterCriteria_DoesNotContain="Không chứa"
                    FilterCriteria_StartsWith="Bắt đầu với" FilterCriteria_EndsWith="Kết thúc với"
                    FilterCriteria_EqualTo="Bằng" FilterCriteria_NotEqualTo="Không bằng" FilterCriteria_SmallerThan="Nhỏ hơn"
                    FilterCriteria_GreaterThan="Lớn hơn" FilterCriteria_SmallerThanOrEqualTo="Nhỏ hơn hoặc bằng"
                    FilterCriteria_GreaterThanOrEqualTo="Lớn hơn hoặc bằng" FilterCriteria_IsNull="Rỗng"
                    FilterCriteria_IsNotNull="Không rỗng" FilterCriteria_IsEmpty="Trống" FilterCriteria_IsNotEmpty="Không trống"
                    Paging_OfText="của" Grouping_GroupingAreaText="" JSWarning="Có một lỗi khởi tạo lưới với ID '[GRID_ID]'. \ N \ n [Chú ý] \ n \ nHãy liên hệ bộ phận bảo trì của Nhất Tâm Soft để được giúp đỡ."
                    LoadingText="Đang tải...." MaxLengthValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX vượt quá số lượng tối đa ký tự YYYYY cho phép cột này."
                    ModifyLink="Chỉnh sửa" NoRecordsText="Không có dữ liệu" Paging_ManualPagingLink="Trang kế »"
                    Paging_PageSizeText="Số dòng:" Paging_PagesText="Trang:" Paging_RecordsText="Dòng:"
                    ResizingTooltipWidth="Rộng:" SaveAllLink="Lưu tất cả" SaveLink="Lưu" TypeValidationError="Giá trị mà bạn đã nhập vào trong cột XXXXX là không đúng."
                    UndeleteLink="Không xóa" UpdateLink="Lưu" />
            </cc3:Grid>
        </fieldset>
    </owd:Window>
    <div>
        <owd:Dialog ID="idDialogYesNo" runat="server" IsModal="true" ShowCloseButton="true"
            Top="0" Left="250" Height="150" Width="300" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
            Title="Cảnh báo">
            <center>
                <br />
                Bạn có đồng ý xóa đối tượng này không?<br />
                Chọn <b>"Đồng ý"</b> để xóa, chọn <b>"Bỏ qua"</b> để đóng!
                <br />
                <br />
                <table>
                    <tr>
                        <td style="width: 85"></td>
                        <td>
                            <asp:Button ID="Button1" runat="server" Width="80" Text="Đồng ý" OnClientClick="xoaDuLieu();return false;" />
                        </td>
                        <td>
                            <asp:Button ID="Button2" runat="server" Text="Bỏ qua" Width="80" OnClientClick="idDialogYesNo.Close();return false;" />
                        </td>
                    </tr>
                </table>
            </center>
        </owd:Dialog>
        <owd:Dialog ID="Dialog2" runat="server" IsModal="true" ShowCloseButton="true" Top="0"
            Left="250" Height="150" Width="300" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
            Title="Cảnh báo">
            <center>
                <br />
                Bạn có đồng ý xóa quyết định khen thưởng này không? Chọn "Đồng ý" để xóa, chọn "Bỏ
                qua" để đóng!
                <br />
                <br />
                <table>
                    <tr>
                        <td style="width: 85"></td>
                        <td>
                            <asp:Button ID="Button3" runat="server" Width="80" Text="Đồng ý" OnClientClick="xoaTTQuyetDinhKT();return false;" />
                        </td>
                        <td>
                            <asp:Button ID="Button4" runat="server" Text="Bỏ qua" Width="80" OnClientClick="Dialog2.Close();return false;" />
                        </td>
                    </tr>
                </table>
            </center>
        </owd:Dialog>
        <owd:Dialog ID="msgDialog" runat="server" IsModal="true" ShowCloseButton="true" Top="0"
            Left="250" Height="150" Width="300" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
            Title="Cảnh báo">
            <center>
                <br />
                <div class="msgDialog">
                </div>
                <br />
                <br />
                <table>
                    <tr>
                        <td>
                            <asp:Button ID="Button6" runat="server" Text="Bỏ qua" Width="80" OnClientClick="msgDialog.Close();return false;" />
                        </td>
                    </tr>
                </table>
            </center>
        </owd:Dialog>
    </div>
    <ajaxToolkit:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
    </ajaxToolkit:ToolkitScriptManager>

    <script>
        var _img = new Array();
        var step = 0;
        function slideit() {
            if (_img.length == 0) {
                step = 0;
                return false;
            }
            $('#anh').fadeIn(3000).attr('src', _img[step]);
        }
        function hinhSau() {
            if (step == 0)
                step = _img.length - 1;
            else
                step = step - 1;;
            slideit();
        }
        function hinhTruoc() {
            if (step == _img.length - 1)
                step = 0;
            else
                step = step + 1;
            slideit();
        }
        function xemAnh(sttTDKT) {
            var i = 0;
            while (_img.length > i) {
                _img[i] = "";
                i++;
            }
            showPopup();
            document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTTDKT").value = sttTDKT;
            var _dtHinh = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.dsHinhAnh(sttTDKT);
            for (var i = 0; i < _dtHinh.value.Rows.length; i++) {
                _img[i] = _dtHinh.value.Rows[i].duongDan;
            }
            slideit();
            return false;
        }
        function xoaAnh() {
            $('#anh').fadeIn(3000).attr('src', "");
            QLNS2014.ThiDuaKhenThuong.nhapktcanhan.xoaFileDinhKem(_img[step]);
            _img = new Array();
            var _dtHinh = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.dsHinhAnh(document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTTDKT").value);
            for (var i = 0; i < _dtHinh.value.Rows.length; i++) {
                _img[i] = _dtHinh.value.Rows[i].duongDan;
            }
            $('#anh').fadeIn(3000).attr('src', _img[_img.length - 1]);
        }
        function uploadStarted() {
        }
        function uploadError(sender, args) {

        }
        function uploadComplete(sender, args) {
            var fileExtension = args.get_fileName();

            if (fileExtension.indexOf('.gif') != -1 || fileExtension.indexOf('.jpg') != -1 || fileExtension.indexOf('.png') != -1 || fileExtension.indexOf('.zip') != -1) {

                _img = new Array();
                var _dtHinh = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.dsHinhAnh(document.getElementById("ctl00_ContentPlaceHolder1_hdfSTTTDKT").value);
                for (var i = 0; i < _dtHinh.value.Rows.length; i++) {
                    _img[i] = _dtHinh.value.Rows[i].duongDan;
                }
                $('#anh').fadeIn(3000).attr('src', _img[_img.length - 1]);
            }
            else {
                alert("Tập tin không đúng định dạng");
            }
            document.getElementById("ctl00_ContentPlaceHolder1_Dialog1_AsyncFileUpload_ctl02").value = "";
            return false;
        }
        function btnNhapExcel_Click() {
            var permiss = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.PhanQuyenChucnang();
            arrValue = permiss.value.split(';');
            if (arrValue[6] == "true") {
                var permiss_ = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KiemTraKhoaChucNang();
                if (permiss_.value == true) {
                    alert("User đang sử dụng đã bị khóa không thực hiện thao tác được!");
                    return false;
                }
            }
            if (arrValue[0] == "true") {
                alert("User đang sử dụng chỉ được xem dữ liệu, không thực hiện được thao tác nhập Excel!");
                return false;
            }
            dlgUpload.screenCenter();
            dlgUpload.Open();
            return false;
        }
        function dongYNhapExcel() {
            DialogYesNoNhapExcel.Close();
            var kq = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.NhapDuLieuFileExcel());
            document.getElementById("info").innerHTML = kq.value;
            for (i = 0; i < Grid1.Rows.length; i++) {
                Grid1.deselectRecord(i);
            }
            Grid1.SelectedRecordsContainer.value = "";
            Grid1.SelectedRecords = new Array();
            Grid1.refresh();
            document.getElementById('processExcel').style.visibility = "hidden";
            return false;
        }
        function btnNhapWord_Click() {
            window.open("/ThiDuaKhenThuong/nhapkhenthuongtheoqd.aspx");
            return false;
        }
        function btnTaiExcekMau_Click() {
            var kq = QLNS2014.ThiDuaKhenThuong.nhapktcanhan.kiemTraTaiExcelMau().value;
            if (kq != "") {
                window.open(kq);
            }
            return false;
        }
        function onClientOpen_Dialog1() {
            document.getElementById("info").innerHTML = "";
            document.getElementById('processExcel').style.visibility = "hidden";
            return false;
        }
        function uploadError_1(sender, args) {

            alert("Tải quyết định thất bại!");
        }
        function uploadComplete_1(sender, args) {
            debugger;
            var kq = (QLNS2014.ThiDuaKhenThuong.nhapktcanhan.KTraFile());
            var fileExtension = args.get_fileName();

            if (fileExtension.indexOf('.xls') != -1) {

                if (kq == false) {
                    alert("Tập tin excel không đúng cấu trúc");
                    return false;
                } else {
                    document.getElementById("info").innerHTML = "";
                    dlgUpload.Close();
                    for (i = 0; i < Grid1.Rows.length; i++) {
                        Grid1.deselectRecord(i);
                    }
                    Grid1.SelectedRecordsContainer.value = "";
                    Grid1.SelectedRecords = new Array();
                    Grid1.refresh();
                }

            } else {
                document.getElementById("info").innerHTML = "<br/>Tập tin không đúng định dạng.";
                return;
            }

        }
    </script>

    <div style="display: none;">
        <ajaxToolkit:AsyncFileUpload OnClientUploadError="uploadError" OnClientUploadComplete="uploadComplete"
            OnClientUploadStarted="uploadStarted" runat="server" ID="AsyncFileUpload" Width="390px"
            UploaderStyle="Traditional" UploadingBackColor="" ThrobberID="myThrobber" BorderStyle="NotSet"
            Font-Underline="False" Font-Strikeout="False" />
    </div>
    <div class="popup" data-popup="popup-1">
        <div class="popup-inner">
            <center>
                <table width="100%">
                    <tr>
                        <td style="width: 5%">
                            <a href="javascript:void(0)" onclick="hinhSau(); return false;">
                                <img src="../images/prev-arrow.png" /></a>
                        </td>
                        <td style="width: 90%">
                            <img id="anh" width="100%" height="550px" />
                        </td>
                        <td style="width: 5%">
                            <a href="javascript:void(0)" onclick="hinhTruoc(); return false;">
                                <img src="../images/next-arrow.png" /></a>
                        </td>
                    </tr>
                </table>
                <br />
                <a href="javascript:void(0)" onclick="document.getElementById('<%=AsyncFileUpload.ClientID%>_ctl02').click();return false;">Thêm ảnh</a> | <a href="javascript:void(0)" onclick="xoaAnh();return false;">Xóa ảnh</a>
            </center>
            <a class="popup-close" data-popup-close="popup-1" href="#">x</a>
        </div>
    </div>
    <style type="text/css">
        .neoslideshow {
            padding: 0px;
            position: relative;
            width: 700px;
            height: 500px;
        }

            .neoslideshow img {
                position: absolute;
                left: 0;
                top: 0;
                z-index: 10;
            }

        #galprev, #galnext {
            position: absolute;
            z-index: 20;
            top: 250px;
            cursor: pointer;
            background: #000;
            color: #fff;
            width: 28px;
            height: 20px;
            line-height: 20px;
            text-align: center;
        }

        #galprev {
            left: 0;
        }

        #galnext {
            right: 0;
        }
        /* Show Popup*/ /* Outer */

        .popup {
            width: 100%;
            height: 100%;
            display: none;
            position: fixed;
            top: 0px;
            left: 0px;
            background: rgba(0,0,0,0.75);
            z-index: 9999999999;
        }
        /* Inner */ .popup-inner {
            max-width: 960px;
            width: 90%;
            padding: 10px 20px 10px 20px;
            position: absolute;
            top: 50%;
            left: 50%;
            -webkit-transform: translate(-50%, -50%);
            transform: translate(-50%, -50%);
            box-shadow: 0px 2px 6px rgba(0,0,0,1);
            border-radius: 3px;
            background: #fff;
            height: 600px;
        }
        /* Close Button */ .popup-close {
            width: 30px;
            height: 30px;
            padding-top: 8px;
            display: inline-block;
            position: absolute;
            top: 10px;
            right: 10px;
            transition: ease 0.25s all;
            -webkit-transform: translate(50%, -50%);
            transform: translate(50%, -50%);
            border-radius: 1000px;
            background: rgba(0,0,0,0.8);
            font-family: Arial, Sans-Serif;
            font-size: 20px;
            text-align: center;
            line-height: 100%;
            color: #fff;
        }

            .popup-close:hover {
                -webkit-transform: translate(50%, -50%) rotate(180deg);
                transform: translate(50%, -50%) rotate(180deg);
                background: rgba(0,0,0,1);
                text-decoration: none;
            }
    </style>
    <owd:Dialog ID="Dialog3" runat="server" IsModal="true" ShowCloseButton="true" Top="0"
        Left="250" Height="300" Width="400" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        Title="Thông báo">
        <div id="chiTietLoi" style="overflow: scroll; width: 390px; height: 250px; margin-top: 5px;">
        </div>
    </owd:Dialog>
    <%--upload file--%>
    <owd:Dialog ID="dlgUpload" runat="server" IsModal="true" ShowCloseButton="true" Top="0"
        Left="250" Height="150" Width="400" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        Title="Chọn tập tin" OnClientOpen="onClientOpen_Dialog1()">
        <div style="padding-top: 10px">
            <cc1:OboutCheckBox ID="chkNhapDauKy" runat="server" Text="Nhập khen thưởng các năm trước"
                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutCheckBox">
            </cc1:OboutCheckBox>
            <ajaxToolkit:AsyncFileUpload OnClientUploadError="uploadExcelError" OnClientUploadComplete="uploadExcelComplete"
                OnClientUploadStarted="uploadExcelStarted" runat="server" ID="AsyncFileUpload2"
                Width="390px" UploaderStyle="Traditional" UploadingBackColor="" ThrobberID="myThrobber"
                BorderStyle="NotSet" Font-Underline="False" Font-Strikeout="False" />
            <br />
            <span><i>Chỉ cho phép tải lên tập tin định dạng: xls,xlsx</i></span><br />
            <span id="info" style="color: Red"></span>
            <div id="processExcel" style="visibility: hidden; display: block;">
                <img alt="" src="/ThiDuaKhenThuong/images/029.gif" width="120px" />
            </div>
        </div>
    </owd:Dialog>
    <owd:Dialog ID="Dialog_ThongBaoLoi" runat="server" IsModal="true" ShowCloseButton="true"
        Top="0" Left="250" Height="250" Width="500" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        Title="Nội dung chi tiết">
        <br />
        <div class="noidungloi" id="noidungloi" style="text-align: left; color: red; overflow: auto; height: 196px;">
        </div>
    </owd:Dialog>
    <owd:Window ID="Window4" runat="server" IsModal="true" ShowCloseButton="true" Status=""
        ShowStatusBar="false" RelativeElementID="WindowPositionHelper" Height="238" Width="550"
        VisibleOnLoad="false" Title="Thêm mới hội đồng Thi đua khen thưởng" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        IsResizable="false">
        <div style="margin-top: 5px">
            <cc1:OboutButton ID="LuuVaDongSKKN" runat="server" Text="Lưu và đóng" Width="100px"
                OnClientClick="LuuvaDongSKKN(); return false;">
            </cc1:OboutButton>
            <cc1:OboutButton ID="DongSKKN" runat="server" Text="Đóng" Width="100px" OnClientClick="Window4.Close(); return false;">
            </cc1:OboutButton>
        </div>
        <div style="margin-top: 5px">
            <fieldset style="border: 1px solid #DBDBE1;">
                <legend><b>Thông tin hội đồng Thi đua khen thưởng</b></legend>
                <table style="width: 100%">
                    <tr>
                        <td style="width: 75px">Tên hội đồng <span style="color: Red">[*]</span>
                        </td>
                        <td colspan="2">
                            <cc1:OboutTextBox ID="txtTenHoiDong" runat="server" Width="400px" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox">
                            </cc1:OboutTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td>Ngày lập <span style="color: Red">[*]</span>
                        </td>
                        <td style="width: 240px">
                            <cc1:OboutTextBox ID="txtNgayLap" runat="server" Width="100%" MaxLength="10" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"></cc1:OboutTextBox>
                        </td>
                        <td style="width: 20px">
                            <obout:Calendar ID="Calendar1" runat="server" Columns="1" DateFormat="dd/MM/yyyy"
                                DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                                DatePickerMode="true" TextBoxId="txtNgayLap" CultureName="vi-VN">
                            </obout:Calendar>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px;">Năm xét duyệt <span style="color: Red">[*]</span>
                        </td>
                        <td style="width: 230px" colspan="2">
                            <cc2:ComboBox ID="cboNamXetDuyet" runat="server" Height="150" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox"
                                FilterType="Contains">
                            </cc2:ComboBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Ghi chú
                        </td>
                        <td style="width: 230px" colspan="2">
                            <cc1:OboutTextBox ID="txtGhiChu" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox">
                            </cc1:OboutTextBox>
                        </td>
                    </tr>
                </table>
            </fieldset>
        </div>
    </owd:Window>
    <owd:Dialog ID="DialogYesNoNhapExcel" runat="server" IsModal="true" ShowCloseButton="false"
        Top="0" Left="250" Height="250" Width="400" VisibleOnLoad="false" StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma"
        Title="Cảnh báo">
        <center>
            <br />
            <div class="noidungkiemtraexcel" id="noidungkiemtraexcel" style="text-align: left; color: red; overflow: auto; height: 160px;">
            </div>
            <table style="width: 100%; text-align: center">
                <tr>
                    <td style="width: 20%"></td>
                    <td style="width: 30%">
                        <cc1:OboutButton ID="Button5" runat="server" Width="80" Text="Đồng ý" OnClientClick="dongYNhapExcel();return false;"
                            FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton" />
                    </td>
                    <td style="width: 30%">
                        <cc1:OboutButton ID="Button7" runat="server" Text="Không đồng ý" Width="100" OnClientClick="DialogYesNoNhapExcel.Close();dlgUpload.Close();return false;"
                            FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutButton" />
                    </td>
                    <td style="width: 20%"></td>
                </tr>
            </table>
        </center>
    </owd:Dialog>
    <owd:Dialog ID="windowThemMoiDoiTuong" runat="server" IsModal="true" ShowCloseButton="false"
        Top="0" Left="250" Height="400" Width="620" zIndex="100000" VisibleOnLoad="false"
        StyleFolder="~/ThiDuaKhenThuong/App_Themes/Styles/wdstyles/dogma" Title="Thêm nhanh đối tượng">
        <div style="margin-top: 5px; margin-bottom: 5px">
            <cc1:OboutButton ID="bttLuuvaDongDoiTuong" runat="server" Text="Lưu và đóng" Width="100px"
                OnClientClick="LuuvaDongDoiTuong(); return false;">
            </cc1:OboutButton>
            <cc1:OboutButton ID="bttDongCN" runat="server" Text="Đóng" Width="100px" OnClientClick="windowThemMoiDoiTuong.Close(); return false;">
            </cc1:OboutButton>
        </div>
        <div>
            <fieldset style="border: 1px solid #DBDBE1;">
                <legend><b>Thông tin cá nhân thi đua, khen thưởng</b></legend>
                <table border="0">
                    <tr>
                        <td style="width: 130px;">Mã CBCC <span style="color: Red">[*]</span>
                        </td>
                        <td colspan="2">
                            <cc1:OboutTextBox ID="txtMaDoiTuong" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox">
                            </cc1:OboutTextBox>
                        </td>
                        <td style="padding-left: 10px;">Đối tượng
                        </td>
                        <td colspan="3">
                            <cc2:ComboBox ID="cboDanhXung" runat="server" Width="100%" Height="100" SelectedIndex="0"
                                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox" AppendDataBoundItems="false">
                                <ClientSideEvents OnItemClick="layGioiTinh" />
                            </cc2:ComboBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Họ và tên <span style="color: Red">[*]</span>
                        </td>
                        <td colspan="5">
                            <cc1:OboutTextBox ID="_txtTenDoiTuong" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox">
                            </cc1:OboutTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Chức vụ hiện tại
                        </td>
                        <td style="width: 180px" colspan="4">
                            <cc1:OboutTextBox ID="txtChucVuHienTai" runat="server" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox">
                            </cc1:OboutTextBox>
                        </td>
                        <td style="float: left; padding-top: 12px;">
                            <cc1:OboutCheckBox ID="chkLaLanhDao" runat="server" Text="Là lãnh đạo" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutCheckBox">
                            </cc1:OboutCheckBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Ngày vào ngành
                        </td>
                        <td colspan="5">
                            <cc1:OboutTextBox ID="txtNgayVeCQ" runat="server" Width="92%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox">
                            </cc1:OboutTextBox>
                            <obout:Calendar ID="Calendar3" runat="server" Columns="1" DateFormat="dd/MM/yyyy"
                                DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                                DatePickerMode="true" TextBoxId="txtNgayVeCQ" CultureName="vi-VN">
                            </obout:Calendar>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Giới tính
                        </td>
                        <td colspan="2">
                            <cc2:ComboBox ID="cboGioiTinh" runat="server" Height="100" SelectedIndex="0" DataSourceID="sdsGioiTinh"
                                FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox" DataTextField="gioiTinh"
                                DataValueField="gioiTinh" AppendDataBoundItems="false">
                            </cc2:ComboBox>
                        </td>
                        <td style="width: 100px; padding-left: 10px">Ngày sinh
                        </td>
                        <td style="width: 180px" colspan="2">
                            <cc1:OboutTextBox ID="_txtNgaySinh" runat="server" Width="80%" MaxLength="10" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutTextBox"></cc1:OboutTextBox>
                            <obout:Calendar ID="Calendar5" runat="server" Columns="1" DateFormat="dd/MM/yyyy"
                                DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                                DatePickerMode="true" TextBoxId="_txtNgaySinh" CultureName="vi-VN">
                            </obout:Calendar>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">CMND
                        </td>
                        <td colspan="2">
                            <cc1:OboutTextBox ID="_txtCMND" runat="server">
                            </cc1:OboutTextBox>
                        </td>
                        <td style="width: 100px; padding-left: 10px">Ngày cấp
                        </td>
                        <td colspan="2">
                            <cc1:OboutTextBox ID="_txtNgayCap" runat="server" Width="80%"></cc1:OboutTextBox>
                            <obout:Calendar ID="Calendar6" runat="server" Columns="1" DateFormat="dd/MM/yyyy"
                                DatePickerImagePath="~/ThiDuaKhenThuong/App_Themes/Styles/Date/date_picker1.gif"
                                DatePickerMode="true" TextBoxId="_txtNgayCap" CultureName="vi-VN">
                            </obout:Calendar>
                        </td>
                    </tr>
                    <tr>
                        <td>Nơi cấp
                        </td>
                        <td colspan="5">
                            <cc1:OboutTextBox Width="100%" ID="_txtNoiCap" runat="server"></cc1:OboutTextBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Phòng ban (tổ bộ môn)
                        </td>
                        <td colspan="5">
                            <cc2:ComboBox ID="cboPhongBan_DT" runat="server" Height="150" FilterType="Contains"
                                AutoClose="true" MenuWidth="425px" Width="100%" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox">
                            </cc2:ComboBox>
                        </td>
                    </tr>
                    <tr>
                        <td style="width: 130px">Đơn vị công tác <span style="color: Red">[*]</span>
                        </td>
                        <td colspan="5">
                            <cc2:ComboBox ID="cbdonViCongTac" runat="server" Height="150" FilterType="Contains"
                                Width="100%" AutoClose="true" MenuWidth="478px" FolderStyle="~/ThiDuaKhenThuong/App_Themes/Styles/Interface/OboutComboBox">
                            </cc2:ComboBox>
                        </td>
                    </tr>
                </table>
            </fieldset>
        </div>
    </owd:Dialog>
    <asp:SqlDataSource ID="sdsGioiTinh" runat="server" SelectCommand="SELECT '' as gioiTinh UNION ALL SELECT N'Nam' as gioiTinh UNION ALL SELECT N'Nữ' as gioiTinh "
        OnLoad="sdsGioiTinh_Load"></asp:SqlDataSource>
</asp:Content>
