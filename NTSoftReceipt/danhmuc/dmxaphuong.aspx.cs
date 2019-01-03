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
using Obout.ComboBox;
using WEB_DLL;

namespace NTSoftReceipt.danhmuc
{
    public partial class dmxaphuong : System.Web.UI.Page
    {
        [AjaxPro.AjaxMethod]
        public string PhanQuyenChucnang()
        {
            string permiss = "";
            string _vPermissValue = "";
            string CurrentFilePath = "";
            try
            {
                CurrentFilePath = Session["CurrentFilePath"].ToString();
                if (CurrentFilePath.ToString() != "")
                {
                    //SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString1());
                    //permiss = sqlFun.GetOneStringField("SELECT permission FROM tblUserPermiss WHERE functionIDpr_sd=(SELECT TOP 1 functionIDpr FROM dbo.tblFunctions WHERE pathFile LIKE N'%" + CurrentFilePath + "%') AND maNguoidungpr_sd=" + HttpContext.Current.Session.GetCurrentUserID() + "");

                    //string _vPermiss = permiss.ToString();
                    //_vPermiss = WEB_DLL.ntsSecurity._mDecrypt(_vPermiss, "rateAnd2012", true).Split(';')[2];
                    //_vPermissValue += ntsSecurityServices.HasPermission(TypeAudit.View, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.AddNew, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.Delete, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.Edit, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.LoadData, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.Print, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.PlusP1, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.PlusP2, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                    //_vPermissValue += ";" + ntsSecurityServices.HasPermission(TypeAudit.PlusP3, Convert.ToInt32(_vPermiss)).ToString().ToLower();
                }
            }
            catch
            { }
            return _vPermissValue;
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            ((CheckBoxColumn)Grid1.Columns[9]).ControlType = GridControlType.Standard;
            if (!IsPostBack)
            {
                Grid1_OnRebind(sender, e);
            }
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmxaphuong));

        }
        protected void OboutCombo_ToolTip(object sender, ComboBoxItemEventArgs e)
        {
            e.Item.ToolTip = e.Item.Text;
        }
        protected void sdsNhomNoiDT_Load(object sender, EventArgs e)
        {
            sdsTinh.ConnectionString = Session.GetConnectionString2();
        }
        protected void sdsNoiDT_Load(object sender, EventArgs e)
        {
            sdsHuyen.ConnectionString = Session.GetConnectionString2();
            sdsHuyen.SelectParameters[0].DefaultValue = "";
        }
        protected void cboNoiDaoTao_LoadingItems(object sender, ComboBoxLoadingItemsEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Text))
            {
                int tryout = 0;
                if (int.TryParse(e.Text, out tryout))
                {
                    sdsHuyen.SelectParameters[0].DefaultValue = e.Text;
                }
            }
        }
        protected void Grid1_OnRebind(object sender, EventArgs e)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("select *,(select tenTinh from tblDMTinh where maTinhpr = maTinhpr_sd) as tenTinh,(SELECT tenHuyen FROM dbo.tblDMQuanHuyen WHERE maHuyenpr=maHuyenpr_sd) AS tenQuanHuyen from tblDMXa");
            Grid1.DataBind();
        }
        [AjaxPro.AjaxMethod]
        public void ThemDuLieu(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand("INSERT INTO dbo.tblDMXa( maXapr ,tenXa ,ghiChu ,ngungSD,tenVietTat,maTinhpr_sd,maHuyenpr_sd)"
                     + " VALUES(@maXapr ,@tenXa ,@ghiChu ,@ngungSD,@tenVietTat,@maTinhpr_sd,@maHuyenpr_sd)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maXapr", _arrayT[0]));
                sqlCom.Parameters.Add(new SqlParameter("@tenXa", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@tenVietTat", _arrayT[4].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr_sd", _arrayT[5].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maHuyenpr_sd", _arrayT[6].ToString()));
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
                SqlCommand sqlCom = new SqlCommand("UPDATE dbo.tblDMXa SET maXapr = @maXapr,tenXa = @tenXa,tenVietTat=@tenVietTat,ghiChu =@ghiChu,ngungSD = @ngungSD,maTinhpr_sd = @maTinhpr_sd,maHuyenpr_sd=@maHuyenpr_sd WHERE maXapr = @maXapr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maXapr", _arrayT[0]));
                sqlCom.Parameters.Add(new SqlParameter("@tenXa", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@tenVietTat", _arrayT[4].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr_sd", _arrayT[5].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maHuyenpr_sd", _arrayT[6].ToString()));

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
                SqlCommand sqlCom = new SqlCommand("delete from dbo.tblDMXa  WHERE maXapr = @maXapr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maXapr", ma.ToString()));
                sqlCom.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch { }
        }
        [AjaxPro.AjaxMethod]
        public bool KTraTonTai(object ma)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string sqlSel = "select * from tblDMXa where maXapr=N'" + ma.ToString() + "'";
            return sqlFun.CheckHasRecord(sqlSel);
        }
        [AjaxPro.AjaxMethod]
        public bool KTraXoa(object ma)
        {
            string strSQL = "SELECT TABLE_NAME tablename FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = N'maXapr_sd'";
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
                            strSQL = "select maXapr_sd from " + _dt.Rows[i][0] + " where maXapr_sd =N'" + ma + "' ";
                        }
                        else
                        {
                            strSQL = strSQL + " UNION ALL select maXapr_sd from " + _dt.Rows[i][0] + " where maXapr_sd =N'" + ma + "' ";
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
