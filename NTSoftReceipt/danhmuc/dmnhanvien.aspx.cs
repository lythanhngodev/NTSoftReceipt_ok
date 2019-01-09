using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Obout.Grid;
using NTSoftReceipt.Class;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using AjaxControlToolkit;
using System.IO;

namespace NTSoftReceipt.danhmuc
{
    public partial class dmnhanvien : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Grid1.DataSource = null;
                Grid1.DataBind();
                LoadCombobox(_cbTinh, "SELECT DISTINCT maTinhpr, tenTinh FROM tblDMTinh t LEFT JOIN tblDMQuanHuyen q ON t.maTinhpr=q.maTinhpr_sd WHERE (t.ngungSD = 0 OR t.ngungSD IS NULL) and (q.maHuyenpr != NULL or q.maHuyenpr!=0) ORDER BY tenTinh ASC", "tenTinh", "maTinhpr");
                LoadCombobox(_cbHuyen, "SELECT DISTINCT maHuyenpr, tenHuyen FROM tblDMQuanHuyen","tenHuyen", "maHuyenpr");
                LoadCombobox(_cbXa, "SELECT DISTINCT maXapr, tenXa FROM tblDMXa", "tenXa", "maXapr");
                LoadCombobox(_cbPhongBan, "SELECT * FROM tblDMPhongBan", "tenPhongBan", "sttPhongBanpr");
            }
            AsyncFileUpload.UploadedComplete += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload_UploadedComplete);
            AsyncFileUpload.UploadedFileError += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload_UploadedFileError);
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmnhanvien));
        }
        #region "Upload excel"
        void AsyncFileUpload_UploadedComplete(object sender, AsyncFileUploadEventArgs e)
        {
            try
            {
                string path = "";
                string urlFile = "";
                string fileName = "";
                string[] arr = e.filename.Split('\\');
                if (arr.Length > 0)
                    fileName = arr[arr.Length - 1];

                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());

                if (!System.IO.Directory.Exists(Server.MapPath("~/Excel/")))
                {
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/Excel/"));
                }
                path = string.Concat(Server.MapPath("~/Excel/" + fileName));

                if (Path.GetExtension(fileName).Contains(".xls") == false)
                {
                    return;
                }

                if (!System.IO.File.Exists(path))
                {
                    //System.IO.File.Delete(path);
                    AsyncFileUpload.SaveAs(path);
                    urlFile = "~/Excel/";
                    // ĐƯỜNG DẪN FILE
                    Session["ssDuongDanFile"] = null;
                    Session["ssDuongDanFile"] = urlFile + fileName;
                }
                else
                {
                    //Xoa file neu da ton tai                   
                    System.IO.File.Delete(path);
                    AsyncFileUpload_UploadedComplete(sender, e);
                }
            }
            catch
            {

            }
            AsyncFileUpload.ClearState();
            AsyncFileUpload.Dispose();
        }
        void AsyncFileUpload_UploadedFileError(object sender, AsyncFileUploadEventArgs e)
        {
        }
        #endregion NHẬP DỮ LIỆU EXCEL

        private void LoadGrid()
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            DataTable dt = new DataTable();
            dt= _sqlfun.GetData("SELECT maNhanVienpr,tenNhanVien,gioiTinh,diaChi,soDienThoai,thuNhap,tenPhongBan,sttPhongBanpr,ghiChu,namSinh,namSinhnam FROM tblDMNhanVien nv LEFT JOIN tblDMPhongBan pb ON nv.sttPhongBanpr_sd = pb.sttPhongBanpr");
            Grid1.DataSource = dt;
            Grid1.DataBind();
        }
        private void LoadCombobox(Obout.ComboBox.ComboBox _cb, string sql,string textField, string valueField)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            _cb.DataSource = _sqlfun.GetData(sql);
            _cb.DataTextField = textField;
            _cb.DataValueField = valueField;
            _cb.DataBind();
        }

        public void LoadGridWhithKey(string key)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                Grid1.DataSource = _sqlfun.GetData("SELECT DISTINCT maNhanVienpr,tenNhanVien,gioiTinh,diaChi,soDienThoai,thuNhap,tenPhongBan,sttPhongBanpr,ghiChu FROM tblDMNhanVien nv LEFT JOIN tblDMPhongBan pb ON nv.sttPhongBanpr_sd = pb.sttPhongBanpr WHERE tenNhanVien like '%" + key + "%'");
                Grid1.DataBind();
            }
            catch
            {
               
            }
        }
        [AjaxPro.AjaxMethod]
        public DataTable LoadCBHuyen(object tinh)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            DataTable dt = new DataTable();
            dt = null;
            try
            {
                dt = _sqlfun.GetData("SELECT maHuyenpr, tenHuyen FROM tblDMQuanHuyen WHERE maTinhpr_sd = '" + tinh.ToString() + "'");
                return dt;
            }
            catch
            {
                return null;
            }
        }
        [AjaxPro.AjaxMethod]
        public DataTable LoadCBXa(object quanhuyen)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            DataTable dt = new DataTable();
            dt = null;
            try
            {
                dt = _sqlfun.GetData("SELECT maXapr, tenXa FROM tblDMXa WHERE maHuyenpr_sd = '" + quanhuyen.ToString() + "'");
                return dt;
            }
            catch
            {
                return null;
            }
        }
        [AjaxPro.AjaxMethod]
        public bool SuaDuLieu(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"UPDATE tblDMNhanVien SET tenNhanVien = @tenNhanVien,gioiTinh = @gioiTinh,diaChi=@diaChi,soDienThoai =@soDienThoai,thuNhap=@thuNhap,sttPhongBanpr_sd = @sttPhongBanpr_sd,ghiChu = @ghiChu WHERE maNhanVienpr = @maNhanVienpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@tenNhanVien", _arrayT[0].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@gioiTinh", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@diaChi", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@soDienThoai", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@thuNhap", (_arrayT[4].ToString().Length==0)?0:decimal.Parse(_arrayT[4].ToString())));
                sqlCom.Parameters.Add(new SqlParameter("@sttPhongBanpr_sd", decimal.Parse(_arrayT[5].ToString())));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[6].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maNhanVienpr", _arrayT[7].ToString()));
                sqlCom.CommandType = CommandType.Text;
                
                if (sqlCom.ExecuteNonQuery() > 0)
                {
                    sqlConn.Close();
                    return true;
                }
            }
            catch (Exception e){
                return false;
            }
            return true;
        }
        [AjaxPro.AjaxMethod]
        public string ThemDuLieu(object[] _arrayT)
        {
            
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"INSERT INTO tblDMNhanVien (maNhanVienpr,tenNhanVien,gioiTinh,diaChi,soDienThoai,thuNhap,sttPhongBanpr_sd,ghiChu) VALUES (@maNhanVienpr,@tenNhanVien,@gioiTinh,@diaChi,@soDienThoai,@thuNhap,@sttPhongBanpr_sd,@ghiChu)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@tenNhanVien", _arrayT[0].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@gioiTinh", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@diaChi", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@soDienThoai", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@thuNhap", (_arrayT[4].ToString().Length == 0) ? 0 : decimal.Parse(_arrayT[4].ToString())));
                sqlCom.Parameters.Add(new SqlParameter("@sttPhongBanpr_sd", decimal.Parse(_arrayT[5].ToString())));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[6].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maNhanVienpr", _arrayT[7].ToString()));
                sqlCom.CommandType = CommandType.Text;

                if (sqlCom.ExecuteNonQuery() > 0)
                {
                    sqlConn.Close();
                    return "";
                }
            }
            catch (Exception e)
            {
                return e.Message;
            }
            return "";
        }
        [AjaxPro.AjaxMethod]
        public bool XoaNhanVien(object chuoi)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"DELETE FROM tblDMNhanVien WHERE maNhanVienpr=@maNhanVienpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maNhanVienpr", chuoi.ToString()));
                sqlCom.CommandType = CommandType.Text;
                if (sqlCom.ExecuteNonQuery() > 0)
                {
                    sqlConn.Close();
                    return true;
                }
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }
        private string TuTang(string chuoichung, string chuoi)
        {
            if (String.IsNullOrEmpty(chuoi))
            {
                return chuoichung + "001";
            }
            chuoi = chuoi.Replace(chuoichung, "");
            long sohientai = (long.Parse(chuoi).ToString().Length>0)?long.Parse(chuoi):0;
            ++sohientai;
            switch (sohientai.ToString().Length)
            {
                case 1:
                    return chuoichung +"00"+sohientai.ToString();
                case 2:
                    return chuoichung + "0" + sohientai.ToString();
                case 3:
                    return chuoichung + sohientai.ToString();
                default:
                    return chuoichung+sohientai.ToString();
            }
        }
        private string TaoMaNhanVienTuDong()
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string ma = _sqlfun.GetOneStringField("SELECT TOP 1 maNhanVienpr FROM tblDMNhanVien Order by maNhanVienpr DESC");
            return TuTang("NV",ma);
        }
        [AjaxPro.AjaxMethod]
        public string TaoMaNhanVienTuDongAjax()
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string ma = _sqlfun.GetOneStringField("SELECT TOP 1 maNhanVienpr FROM tblDMNhanVien Order by maNhanVienpr DESC");
            return TuTang("NV", ma);
        }
        protected void Grid1_Rebind(object sender, EventArgs e)
        {
            switch (_txtLoadRefesh.Value)
            {
                case "lammoi":
                    LoadGrid();
                    break;
                case "timkiem":
                    LoadGridWhithKey(_txtKey.Value);
                    break;
                default:
                    break;
            }
        }
        [AjaxPro.AjaxMethod]
        public string[] NhapDuLieuExcel()
        {
            string path = Session["ssDuongDanFile"].ToString();
            string loinhanvien = "";
            decimal sodongthemthanhcong = 0;
            decimal sodongtong = 0;
            string[] loi = new string[4];
            // kết nối csdl
            SqlFunction _class = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string ExcelConnectionString = "";
            try
            {
                path = Server.MapPath(path);
                string[] kT = path.Split('.');
                DataTable dtTemp = new DataTable();
                ExcelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'", path);
                OleDbConnection connection = new OleDbConnection(ExcelConnectionString);
                connection.Open();
                string sqlSelect = "select * from [Sheet1$A1:I20000]";
                OleDbCommand conmand = new OleDbCommand(sqlSelect, connection);
                OleDbDataAdapter oleDA = new OleDbDataAdapter(conmand);
                oleDA.Fill(dtTemp);
                connection.Close();
                /*
                foreach (DataRow item in dtTemp.Rows)
                {
                    
                }
                */
                // duyệt từng dòng dữ liệu

                DataTable dtPhongBan = _class.GetData(@"SELECT sttPhongBanpr, tenPhongBan FROM tblDMPhongBan");

                for (int i = 0; i < dtTemp.Rows.Count; i++)
                {
                    string tenNhanVien = "", maNhanVienpr = "", sttPhongBanpr_sd = null, namSinh = "", gioiTinh = "", diaChi = "", soDienThoai = "";
                    decimal thuNhap = 0;
                    bool nhapExcel=false;
                    // kiểm tra tên nhân viên
                    if (!(String.IsNullOrEmpty(dtTemp.Rows[i][2].ToString())))
                    {
                        tenNhanVien = dtTemp.Rows[i][2].ToString();
                        // cộng 1 cho tổng số dòng
                        sodongtong++;
                        // kiểm tra họ tên có hợp lệ hay không
                        bool loie=false;
                        string[] dauDacBiet = { ".", "@", "#", "$", "%", "^", "&", "*", "=" };
                        for (int y = 0; y < dauDacBiet.Length; y++)
                        {
                            if (tenNhanVien.IndexOf(dauDacBiet[y].ToString()) >= 0)
                            {
                                loie = true;
                                break;
                            }
                        }
                        if (loie)
                        {
                            loinhanvien+=" - Nhân viên "+tenNhanVien+" có họ và tên chứa ký tự không hợp lệ.</br>";
                            continue; // xét người tiếp theo
                        }
                    }

                    // kiểm tra mã nhân viên
                    if (!(String.IsNullOrEmpty(dtTemp.Rows[i][1].ToString())))
                    {
                        maNhanVienpr=dtTemp.Rows[i][1].ToString();
                        if (_class.CheckHasRecord(@"SELECT * FROM tblDMNhanVien WHERE maNhanVienpr = '"+ maNhanVienpr + "'"))
                        {
                            loinhanvien += "- Mã nhân viên " + maNhanVienpr + " bị trùng.</br>";
                            continue;
                        }
                    }else
                    {
                        loinhanvien += "- Nhân viên " + tenNhanVien + " chưa có mã nhân viên.</br>";
                        continue;
                    }
                    
                    // kiểm tra phòng ban
                    if (String.IsNullOrEmpty(dtTemp.Rows[i][3].ToString()))
                    {
                        sttPhongBanpr_sd = DBNull.Value.ToString();
                    }
                    else
                    {
                        sttPhongBanpr_sd = dtTemp.Rows[i][3].ToString();
                        bool ck = false;
                        string temp = "";
                        for (int p = 0; p < dtPhongBan.Rows.Count; p++)
                        {
                            if (dtPhongBan.Rows[p][1].ToString() == sttPhongBanpr_sd)
                            {
                                ck = true;
                                temp = dtPhongBan.Rows[p][0].ToString();
                                break;
                            }
                        }
                        if (ck)
                        {
                            sttPhongBanpr_sd = temp;
                        }
                        else
                        {
                            sttPhongBanpr_sd = DBNull.Value.ToString();
                        }
                    }
                    // Kiểm tra năm sinh
                    namSinh = dtTemp.Rows[i][4].ToString();
                    object[] ktNamSinh = KiemTraNamSinh(namSinh);
                    namSinh = ktNamSinh[0].ToString();
                    bool namSinhnam = bool.Parse(ktNamSinh[1].ToString());
                    // Kiểm tra giới tính 
                    gioiTinh = dtTemp.Rows[i][5].ToString();
                    // Kiểm tra địa chỉ
                    diaChi = dtTemp.Rows[i][6].ToString();
                    // Kiểm tra số điện thoại
                    soDienThoai = dtTemp.Rows[i][7].ToString();
                    // kiểm tra mức lương
                    try
                    {
                        thuNhap = decimal.Parse(dtTemp.Rows[i][8].ToString());
                    }
                    catch
                    {
                        loinhanvien += "- Nhân viên "+tenNhanVien+" có số mức lương không đúng.<br/>";
                    }
                    // nhập excel
                    nhapExcel = true;

                    // Thực thi sql
                    string sql = @"
                        INSERT INTO tblDMNhanVien
                                (
                                maNhanVienpr
                                ,tenNhanVien
                                ,sttPhongBanpr_sd
                                ,namSinh
                                ,namSinhnam
                                ,gioiTinh
                                ,diaChi
                                ,soDienThoai
                                ,thuNhap
                                ,nhapExcel
                                )
                            VALUES
                                (
                                @maNhanVienpr
                                ,@tenNhanVien
                                ,@sttPhongBanpr_sd
                                ,@namSinh
                                ,@namSinhnam
                                ,@gioiTinh
                                ,@diaChi
                                ,@soDienThoai
                                ,@thuNhap
                                ,@nhapExcel
                                )
                    ";
                    SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                    sqlConn.Open();
                    SqlCommand sqlCom = new SqlCommand(sql, sqlConn);
                    sqlCom.Parameters.AddWithValue("@maNhanVienpr",maNhanVienpr);
                    sqlCom.Parameters.AddWithValue("@tenNhanVien", tenNhanVien);
                    sqlCom.Parameters.AddWithValue("@sttPhongBanpr_sd", sttPhongBanpr_sd);
                    sqlCom.Parameters.AddWithValue("@namSinh", namSinh);
                    sqlCom.Parameters.AddWithValue("@namSinhnam", namSinhnam);
                    sqlCom.Parameters.AddWithValue("@gioiTinh", gioiTinh);
                    sqlCom.Parameters.AddWithValue("@diaChi", diaChi);
                    sqlCom.Parameters.AddWithValue("@soDienThoai", soDienThoai);
                    sqlCom.Parameters.AddWithValue("@thuNhap", thuNhap);
                    sqlCom.Parameters.AddWithValue("@nhapExcel", nhapExcel);
                    sqlCom.CommandType = CommandType.Text;
                    if (sqlCom.ExecuteNonQuery()!=0)
                        sodongthemthanhcong++;
                }
            }
            catch
            {
                loi[0] = loinhanvien;
                loi[1] = sodongthemthanhcong.ToString();
                loi[2] = (sodongtong - sodongthemthanhcong).ToString();
                loi[3] = sodongtong.ToString();
                return loi;
            }
            loi[0] = loinhanvien;
            loi[1] = sodongthemthanhcong.ToString();
            loi[2] = (sodongtong - sodongthemthanhcong).ToString();
            loi[3] = sodongtong.ToString();
            return loi;
        }
        [AjaxPro.AjaxMethod]
        public string[] NhapDuLieuExcel_Xoa()
        {
            string path = Session["ssDuongDanFile"].ToString();
            string loinhanvien = "";
            decimal sodongthemthanhcong = 0;
            decimal sodongtong = 0;
            string[] loi = new string[4];
            // kết nối csdl
            SqlFunction _class = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            // Xóa tất cả nhân viên
            _class.ExeCuteNonQuery(@"DELETE FROM tblDMNhanVien WHERE 1=1");
            string ExcelConnectionString = "";
            try
            {
                path = Server.MapPath(path);
                string[] kT = path.Split('.');
                DataTable dtTemp = new DataTable();
                ExcelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1'", path);
                OleDbConnection connection = new OleDbConnection(ExcelConnectionString);
                connection.Open();
                string sqlSelect = "select * from [Sheet1$A1:I20000]";
                OleDbCommand conmand = new OleDbCommand(sqlSelect, connection);
                OleDbDataAdapter oleDA = new OleDbDataAdapter(conmand);
                oleDA.Fill(dtTemp);
                connection.Close();
                /*
                foreach (DataRow item in dtTemp.Rows)
                {
                    
                }
                */
                // duyệt từng dòng dữ liệu

                DataTable dtPhongBan = _class.GetData(@"SELECT sttPhongBanpr, tenPhongBan FROM tblDMPhongBan");

                for (int i = 0; i < dtTemp.Rows.Count; i++)
                {
                    string tenNhanVien = "", maNhanVienpr = "", sttPhongBanpr_sd = null, namSinh = "", gioiTinh = "", diaChi = "", soDienThoai = "";
                    decimal thuNhap = 0;
                    bool nhapExcel = false;
                    // kiểm tra tên nhân viên
                    if (!(String.IsNullOrEmpty(dtTemp.Rows[i][2].ToString())))
                    {
                        tenNhanVien = dtTemp.Rows[i][2].ToString();
                        // cộng 1 cho tổng số dòng
                        sodongtong++;
                        // kiểm tra họ tên có hợp lệ hay không
                        bool loie = false;
                        string[] dauDacBiet = { ".", "@", "#", "$", "%", "^", "&", "*", "=" };
                        for (int y = 0; y < dauDacBiet.Length; y++)
                        {
                            if (tenNhanVien.IndexOf(dauDacBiet[y].ToString()) >= 0)
                            {
                                loie = true;
                                break;
                            }
                        }
                        if (loie)
                        {
                            loinhanvien += " - Nhân viên " + tenNhanVien + " có họ và tên chứa ký tự không hợp lệ.</br>";
                            continue; // xét người tiếp theo
                        }
                    }

                    // kiểm tra mã nhân viên
                    if (!(String.IsNullOrEmpty(dtTemp.Rows[i][1].ToString())))
                    {
                        maNhanVienpr = dtTemp.Rows[i][1].ToString();
                        if (_class.CheckHasRecord(@"SELECT * FROM tblDMNhanVien WHERE maNhanVienpr = '" + maNhanVienpr + "'"))
                        {
                            loinhanvien += "- Mã nhân viên " + maNhanVienpr + " bị trùng.</br>";
                            continue;
                        }
                    }
                    else
                    {
                        loinhanvien += "- Nhân viên " + tenNhanVien + " chưa có mã nhân viên.</br>";
                        continue;
                    }

                    // kiểm tra phòng ban
                    if (String.IsNullOrEmpty(dtTemp.Rows[i][3].ToString()))
                    {
                        sttPhongBanpr_sd = DBNull.Value.ToString();
                    }
                    else
                    {
                        sttPhongBanpr_sd = dtTemp.Rows[i][3].ToString();
                        bool ck = false;
                        string temp = "";
                        for (int p = 0; p < dtPhongBan.Rows.Count; p++)
                        {
                            if (dtPhongBan.Rows[p][1].ToString() == sttPhongBanpr_sd)
                            {
                                ck = true;
                                temp = dtPhongBan.Rows[p][0].ToString();
                                break;
                            }
                        }
                        if (ck)
                        {
                            sttPhongBanpr_sd = temp;
                        }
                        else
                        {
                            sttPhongBanpr_sd = DBNull.Value.ToString();
                        }
                    }
                    // Kiểm tra năm sinh
                    namSinh = dtTemp.Rows[i][4].ToString();
                    object[] ktNamSinh = KiemTraNamSinh(namSinh);
                    namSinh = ktNamSinh[0].ToString();
                    bool namSinhnam = bool.Parse(ktNamSinh[1].ToString());
                    // Kiểm tra giới tính 
                    gioiTinh = dtTemp.Rows[i][5].ToString();
                    // Kiểm tra địa chỉ
                    diaChi = dtTemp.Rows[i][6].ToString();
                    // Kiểm tra số điện thoại
                    soDienThoai = dtTemp.Rows[i][7].ToString();
                    // kiểm tra mức lương
                    try
                    {
                        thuNhap = decimal.Parse(dtTemp.Rows[i][8].ToString());
                    }
                    catch
                    {
                        loinhanvien += "- Nhân viên " + tenNhanVien + " có số mức lương không đúng.<br/>";
                    }
                    // nhập excel
                    nhapExcel = true;

                    // Thực thi sql
                    string sql = @"
                        INSERT INTO tblDMNhanVien
                                (
                                maNhanVienpr
                                ,tenNhanVien
                                ,sttPhongBanpr_sd
                                ,namSinh
                                ,namSinhnam
                                ,gioiTinh
                                ,diaChi
                                ,soDienThoai
                                ,thuNhap
                                ,nhapExcel
                                )
                            VALUES
                                (
                                @maNhanVienpr
                                ,@tenNhanVien
                                ,@sttPhongBanpr_sd
                                ,@namSinh
                                ,@namSinhnam
                                ,@gioiTinh
                                ,@diaChi
                                ,@soDienThoai
                                ,@thuNhap
                                ,@nhapExcel
                                )
                    ";
                    SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                    sqlConn.Open();
                    SqlCommand sqlCom = new SqlCommand(sql, sqlConn);
                    sqlCom.Parameters.AddWithValue("@maNhanVienpr", maNhanVienpr);
                    sqlCom.Parameters.AddWithValue("@tenNhanVien", tenNhanVien);
                    sqlCom.Parameters.AddWithValue("@sttPhongBanpr_sd", sttPhongBanpr_sd);
                    sqlCom.Parameters.AddWithValue("@namSinh", namSinh);
                    sqlCom.Parameters.AddWithValue("@namSinhnam", namSinhnam);
                    sqlCom.Parameters.AddWithValue("@gioiTinh", gioiTinh);
                    sqlCom.Parameters.AddWithValue("@diaChi", diaChi);
                    sqlCom.Parameters.AddWithValue("@soDienThoai", soDienThoai);
                    sqlCom.Parameters.AddWithValue("@thuNhap", thuNhap);
                    sqlCom.Parameters.AddWithValue("@nhapExcel", nhapExcel);
                    sqlCom.CommandType = CommandType.Text;
                    if (sqlCom.ExecuteNonQuery() != 0)
                        sodongthemthanhcong++;
                }
            }
            catch (Exception ex)
            {
                loi[0] = loinhanvien;
                loi[1] = sodongthemthanhcong.ToString();
                loi[2] = (sodongtong - sodongthemthanhcong).ToString();
                loi[3] = sodongtong.ToString();
                return loi;
            }
            loi[0] = loinhanvien;
            loi[1] = sodongthemthanhcong.ToString();
            loi[2] = (sodongtong - sodongthemthanhcong).ToString();
            loi[3] = sodongtong.ToString();
            return loi;
        }
        private object[] KiemTraNamSinh(string chuoi)
        {
            try
            {
                switch (chuoi.Length)
                {
                    case 0:
                        break;
                    case 4:
                        return new object[] { chuoi+"-01-01", true };
                    default:
                        return new object[] { chuoi, false };
                }
            }
            catch
            {
                return new object[] { "", false };
            }
            return new object[] { "", false };
        }
    }
}