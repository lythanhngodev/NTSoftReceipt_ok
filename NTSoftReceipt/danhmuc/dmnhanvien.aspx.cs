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
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmnhanvien));
        }
        private void LoadGrid()
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("SELECT maNhanVienpr,tenNhanVien,gioiTinh,diaChi,soDienThoai,thuNhap,tenPhongBan,sttPhongBanpr,ghiChu FROM tblDMNhanVien nv LEFT JOIN tblDMPhongBan pb ON nv.sttPhongBanpr_sd = pb.sttPhongBanpr");
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

    }
}