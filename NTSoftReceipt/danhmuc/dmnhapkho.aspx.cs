using NTSoftReceipt.Class;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace NTSoftReceipt.danhmuc
{
    public partial class dmnhapkho : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Grid1.DataSource = null;
                Grid1.DataBind();
                Grid2.DataSource = null;
                Grid2.DataBind();
                LoadCombobox(cbNguoiNhap, "SELECT maNhanVienpr, tenNhanVien from tblDMNhanVien", "tenNhanVien", "maNhanVienpr");
                LoadCombobox(cbNhaCC, "SELECT maNhaCCpr, tenNhaCC FROM tblDMNhaCC", "tenNhaCC", "maNhaCCpr");
                LoadCombobox(cbBienLaiW2, "SELECT sttBienLaipr, tenBienLai FROM tblDMBienLai", "tenBienLai", "sttBienLaipr");
            }
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmnhapkho));
        }
        public bool KTraXoa(string ma, string khoa)
        {
            khoa += "_sd";
            string strSQL = "SELECT TABLE_NAME tablename FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = N'"+khoa+"'";
            SqlFunction _sqlClass = new SqlFunction(Session.GetConnectionString2());
            DataTable _dt = new DataTable();
            _dt = _sqlClass.GetData(strSQL);
            strSQL = " ";
            if (_dt.Rows.Count > 0)
            {
                try
                {
                    for (int i = 0; i < _dt.Rows.Count; i++)
                    {
                        if (i == 0)
                        {
                            strSQL = "select "+khoa+" from " + _dt.Rows[i][0] + " where "+khoa+" =N'" + ma + "' ";
                        }
                        else
                        {
                            strSQL = strSQL + " UNION ALL select "+khoa+" from " + _dt.Rows[i][0] + " where "+khoa+" =N'" + ma + "' ";
                        }
                    }
                    return _sqlClass.CheckHasRecord(strSQL);
                }
                catch { return true; }
            }
            else return false;
        }
        [AjaxPro.AjaxMethod]
        public bool SuaNhapKho(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"UPDATE tblNhapKho SET ngayNhap = @ngayNhap,maNhanVienpr_sd = @maNhanVienpr_sd, maNhaCCpr_sd = @maNhaCCpr_sd, noiDung = @noiDung, ghiChu = @ghiChu WHERE sttNKpr = @sttNKpr", sqlConn);
                try
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngayNhap", DateTime.Parse(_arrayT[1].ToString())));
                }
                catch
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngayNhap", DBNull.Value));
                }
                sqlCom.Parameters.Add(new SqlParameter("@maNhanVienpr_sd", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maNhaCCpr_sd", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@noiDung", _arrayT[4].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[5].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@sttNKpr", _arrayT[0].ToString()));
                sqlCom.CommandType = CommandType.Text;
                if (sqlCom.ExecuteNonQuery() != 0)
                {
                    sqlConn.Close();
                    return true;
                }
            }
            catch
            {
                sqlConn.Close();
                return false;
            }
            sqlConn.Close();
            return false;
        }
        [AjaxPro.AjaxMethod]
        public bool LuuCapNhatNhapKhoCT(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"UPDATE tblNhapKhoCT SET sttBienLaipr_sd=@sttBienLaipr_sd,soLuong=@soLuong,donGia=@donGia,thanhTien=@thanhTien WHERE sttNKCTpr=@sttNKCTpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@sttBienLaipr_sd", _arrayT[0].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@soLuong", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@donGia", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@thanhTien", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@sttNKCTpr", _arrayT[4].ToString()));
                
                sqlCom.CommandType = CommandType.Text;
                if (sqlCom.ExecuteNonQuery() != 0)
                {
                    sqlConn.Close();
                    return true;
                }
            }
            catch
            {
                sqlConn.Close();
                return false;
            }
            sqlConn.Close();
            return false;
        }
        [AjaxPro.AjaxMethod]
        public bool ThemNhapKhoCT(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"INSERT INTO tblNhapKhoCT (sttBienLaipr_sd,soLuong,donGia,thanhTien,sttNKpr_sd) VALUES (@sttBienLaipr_sd,@soLuong,@donGia,@thanhTien,@sttNKpr_sd)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@sttBienLaipr_sd", _arrayT[0].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@soLuong", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@donGia", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@thanhTien", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@sttNKpr_sd", _arrayT[4].ToString()));
                sqlCom.CommandType = CommandType.Text;
                if (sqlCom.ExecuteNonQuery() != 0)
                {
                    sqlConn.Close();
                    return true;
                }
            }
            catch
            {
                sqlConn.Close();
                return false;
            }
            sqlConn.Close();
            return false;
        }
        [AjaxPro.AjaxMethod]
        public bool XoaNhapKhoCT(object chuoi)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"DELETE FROM tblNhapKhoCT WHERE sttNKCTpr=@sttNKCTpr", sqlConn);
                sqlCom.Parameters.AddWithValue("@sttNKCTpr", chuoi.ToString());
                sqlCom.CommandType = CommandType.Text;
                int l = sqlCom.ExecuteNonQuery();
                if ( l!= 0)
                {
                    sqlConn.Close();
                    return true;
                }
            }
            catch
            {
                sqlConn.Close();
                return false;
            }
            sqlConn.Close();
            return false;
        }
        [AjaxPro.AjaxMethod]
        public bool XoaNhapKho(object chuoi)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                //public string KiemTraTruocKhiXoa(string sMa, string sChuoiTru, string sCot);
                SqlCommand sqlCom = new SqlCommand(@"DELETE FROM tblNhapKho WHERE sttNKpr=@sttNKpr", sqlConn);
                sqlCom.Parameters.AddWithValue("@sttNKpr", chuoi.ToString());
                sqlCom.CommandType = CommandType.Text;
                if (KTraXoa(chuoi.ToString(), "sttNKpr"))
                {
                    // nếu check có thì trả về false
                    return false;
                }else
                {
                    if (sqlCom.ExecuteNonQuery() != 0)
                    {
                        sqlConn.Close();
                        return true;
                    }
                }
            }
            catch
            {
                sqlConn.Close();
                return false;
            }
            sqlConn.Close();
            return false;
        }
        [AjaxPro.AjaxMethod]
        public int ThemNhapKho(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            int id = 0;
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand(@"INSERT INTO tblNhapKho (soPhieu,ngayNhap,maNhanVienpr_sd,maNhaCCpr_sd,noiDung,ghiChu) VALUES (@soPhieu,@ngayNhap,@maNhanVienpr_sd,@maNhaCCpr_sd,@noiDung,@ghiChu)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@soPhieu", _arrayT[0].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngayNhap", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maNhanVienpr_sd", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maNhaCCpr_sd", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@noiDung", _arrayT[4].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[5].ToString()));
                sqlCom.CommandType = CommandType.Text;
                id = (int)sqlCom.ExecuteScalar();
                if ( id!= 0)
                {
                    sqlConn.Close();
                    return id;
                }
            }
            catch
            {
                sqlConn.Close();
                return id;
            }
            sqlConn.Close();
            return id;
        }
        private void LoadGrid(Obout.Grid.Grid grid, string chuoi)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            grid.DataSource = _sqlfun.GetData(chuoi);
            grid.DataBind();
        }
        private void LoadCombobox(Obout.ComboBox.ComboBox _cb, string sql, string textField, string valueField)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            _cb.DataSource = _sqlfun.GetData(sql);
            _cb.DataTextField = textField;
            _cb.DataValueField = valueField;
            _cb.DataBind();
        }
        protected void Grid1_Rebind1(object sender, EventArgs e)
        {
            switch (txtHDLoaiG1.Value)
            {
                case "lammoi":
                    LoadGrid(Grid1, "SELECT nk.sttNKpr, nk.soPhieu, nk.ngayNhap,ncc.maNhaCCpr,ncc.tenNhaCC,nk.maNhanVienpr_sd, nv.tenNhanVien, nk.noiDung, nk.ghiChu FROM tblNhapKho nk LEFT JOIN tblDMNhanVien nv ON nk.maNhanVienpr_sd = nv.maNhanVienpr LEFT JOIN tblDMNhaCC ncc ON nk.maNhaCCpr_sd = ncc.maNhaCCpr");
                    break;
                default:
                    break;
            }
        }

        protected void Grid2_Rebind(object sender, EventArgs e)
        {
            switch (txtHDLoaiG2.Value)
            {
                case "lammoi":
                    Grid2.DataSource = null;
                    Grid2.DataBind();
                    break;
                case "loadgridtuSttNK":
                    LoadGrid(Grid2, "SELECT ct.sttNKCTpr, ct.sttBienLaipr_sd, bl.tenBienLai, ct.soLuong,ct.donGia,ct.thanhTien FROM tblNhapKho nk, tblNhapKhoCT ct, tblDMBienLai bl WHERE nk.sttNKpr = ct.sttNKpr_sd and ct.sttBienLaipr_sd = bl.sttBienLaipr and sttNKpr_sd = '" + long.Parse(txtSttNKpr_sd.Value)+"'");
                    break;
                default:
                    Grid2.DataSource = null;
                    Grid2.DataBind();
                    break;
            }
        }
    }
}