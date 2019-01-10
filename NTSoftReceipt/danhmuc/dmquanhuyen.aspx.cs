using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Xml.Linq;
using Obout.Grid;
using NTSoftReceipt.Class;
using System.Data.SqlClient;
using WEB_DLL;

namespace NTSoftReceipt.danhmuc
{
    public partial class dmquanhuyen : System.Web.UI.Page
    {
        [AjaxPro.AjaxMethod]
        public string PhanQuyenChucnang()
        {
            string permiss = "";
            string _vPermissValue = "";
            string CurrentFilePath = "";
            //try
            //{
            //    CurrentFilePath = Session["CurrentFilePath"].ToString();
            //    if (CurrentFilePath.ToString() != "")
            //    {
            //        SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString1());
            //        permiss = sqlFun.GetOneStringField("SELECT permission FROM tblUserPermiss WHERE functionIDpr_sd=(SELECT TOP 1 functionIDpr FROM dbo.tblFunctions WHERE pathFile LIKE N'%" + CurrentFilePath + "%') AND maNguoidungpr_sd=" + HttpContext.Current.Session.GetCurrentUserID() + "");

            //        string _vPermiss = permiss.ToString();
            //        _vPermiss = WEB_DLL.ntsSecurity._mDecrypt(_vPermiss, "rateAnd2012", true).Split(';')[2];
            //        _vPermissValue += ntsSecurityServices.HasPermission(TypeAudit.View, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.AddNew, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.Delete, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.Edit, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.LoadData, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.Print, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.PlusP1, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.PlusP2, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //        _vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.PlusP3, Convert.ToInt32(_vPermiss)).ToString().ToLower();
            //    }
            //}
            //catch
            //{ }
            return _vPermissValue;
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            ((CheckBoxColumn)Grid1.Columns[7]).ControlType = GridControlType.Standard;
            if (!IsPostBack)
            {
                Grid1_OnRebind(sender, e);
                //Load combobox
                SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                // trả về bảng
                _cboTinh.DataSource = _sqlfun.GetData("SELECT maTinhpr, tenTinh FROM tblDMTinh WHERE ngungSD = 0 OR ngungSD IS NULL ORDER BY tenTinh ASC");
                _cboTinh.DataValueField = "maTinhpr";
                _cboTinh.DataTextField = "tenTinh";
                _cboTinh.DataBind();
            }
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmquanhuyen));

        }
        protected void Grid1_OnRebind(object sender, EventArgs e)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("select *,(select tenTinh from tblDMTinh where maTinhpr = maTinhpr_sd) as tenTinh from tblDMQuanHuyen");
            Grid1.DataBind();
        }
        [AjaxPro.AjaxMethod]
        public void ThemDuLieu(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand("INSERT INTO dbo.tblDMQuanHuyen( maHuyenpr ,tenHuyen ,ghiChu ,ngungSD,tenVietTat,maTinhpr_sd)"
                     + " VALUES(@maHuyenpr ,@tenHuyen ,@ghiChu ,@ngungSD,@tenVietTat,@maTinhpr_sd)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maHuyenpr", _arrayT[0]));
                sqlCom.Parameters.Add(new SqlParameter("@tenHuyen", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@tenVietTat", _arrayT[4].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr_sd", _arrayT[5].ToString()));
                sqlCom.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch { }
        }
        [AjaxPro.AjaxMethod]
        public void SuaDuLieu(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand("UPDATE dbo.tblDMQuanHuyen SET maHuyenpr = @maHuyenpr,tenHuyen = @tenHuyen,tenVietTat=@tenVietTat,ghiChu =@ghiChu,ngungSD = @ngungSD,maTinhpr_sd = @maTinhpr_sd WHERE maHuyenpr = @maHuyenpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maHuyenpr", _arrayT[0]));
                sqlCom.Parameters.Add(new SqlParameter("@tenHuyen", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@tenVietTat", _arrayT[4].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr_sd", _arrayT[5].ToString()));
                sqlCom.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch { }
        }
        [AjaxPro.AjaxMethod]
        public void XoaDuLieu(object ma)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand("delete from dbo.tblDMQuanHuyen  WHERE maHuyenpr = @maHuyenpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maHuyenpr", ma.ToString()));
                sqlCom.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch { }
        }
        [AjaxPro.AjaxMethod]
        public bool KTraTonTai(object ma)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string sqlSel = "select * from tblDMQuanHuyen where maHuyenpr=N'" + ma.ToString() + "'";
            return sqlFun.CheckHasRecord(sqlSel);
        }
        [AjaxPro.AjaxMethod]
        public bool KTraXoa(object ma)
        {
            string strSQL = "SELECT TABLE_NAME tablename FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = N'maHuyenpr_sd'";
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
                            strSQL = "select maHuyenpr_sd from " + _dt.Rows[i][0] + " where maHuyenpr_sd =N'" + ma + "' ";
                        }
                        else
                        {
                            strSQL = strSQL + " UNION ALL select maHuyenpr_sd from " + _dt.Rows[i][0] + " where maHuyenpr_sd =N'" + ma + "' ";
                        }
                    }
                    return _sqlClass.CheckHasRecord(strSQL);
                }
                catch { return true; }
            }
            else return false;
        }
    }
}
