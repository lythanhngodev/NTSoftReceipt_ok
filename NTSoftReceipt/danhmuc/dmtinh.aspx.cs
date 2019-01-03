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
    public partial class dmtinh : System.Web.UI.Page
    {
        [AjaxPro.AjaxMethod]
        protected void Page_Load(object sender, EventArgs e)
        {
            //dinh style cho checkbox
            ((CheckBoxColumn)Grid1.Columns[5]).ControlType = GridControlType.Standard;
            //chỉ chạy lần đầu tiên
            if (!IsPostBack)
            {
                Grid1_OnRebind(sender, e);
            }
            //đăng ký AjaxPro để gọi hàm từ javascript qua C#
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmtinh));

        }
        protected void Grid1_OnRebind(object sender, EventArgs e)
        {
            //HttpContext.Current.Session.GetConnectionString2() chứa kết nối đến database đang thao tác
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("select * from tblDMTinh");
            Grid1.DataBind();
        }
        [AjaxPro.AjaxMethod]//khai báo để gọi từ javascript
        public void ThemDuLieu(object[] _arrayT)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand("INSERT INTO dbo.tblDMTinh( maTinhpr ,tenTinh ,ghiChu ,ngungSD,tenVietTat)"
                     + " VALUES(@maTinhpr ,@tenTinh ,@ghiChu ,@ngungSD,@tenVietTat)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr", _arrayT[0]));
                sqlCom.Parameters.Add(new SqlParameter("@tenTinh", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@tenVietTat", _arrayT[4].ToString()));
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
                SqlCommand sqlCom = new SqlCommand("UPDATE dbo.tblDMTinh SET maTinhpr = @maTinhpr,tenTinh = @tenTinh,tenVietTat=@tenVietTat,ghiChu =@ghiChu,ngungSD = @ngungSD WHERE maTinhpr = @maTinhpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr", _arrayT[0]));
                sqlCom.Parameters.Add(new SqlParameter("@tenTinh", _arrayT[1].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ghiChu", _arrayT[2].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", _arrayT[3].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@tenVietTat", _arrayT[4].ToString()));
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
                SqlCommand sqlCom = new SqlCommand("delete from dbo.tblDMTinh  WHERE maTinhpr = @maTinhpr", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@maTinhpr", ma.ToString()));
                sqlCom.ExecuteNonQuery();
                sqlConn.Close();
            }
            catch { }
        }
        [AjaxPro.AjaxMethod]
        public bool KTraTonTai(object ma)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string sqlSel = "select * from tblDMTinh where maTinhpr=N'" + ma.ToString() + "'";
            return sqlFun.CheckHasRecord(sqlSel);
        }
        [AjaxPro.AjaxMethod]
        public bool KTraXoa(object ma)
        {
            string strSQL = "SELECT TABLE_NAME tablename FROM INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME = N'maTinhpr_sd'";
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
                            strSQL = "select maTinhpr_sd from " + _dt.Rows[i][0] + " where maTinhpr_sd =N'" + ma + "' ";
                        }
                        else
                        {
                            strSQL = strSQL + " UNION ALL select maTinhpr_sd from " + _dt.Rows[i][0] + " where maTinhpr_sd =N'" + ma + "' ";
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
