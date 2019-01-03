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
using Obout.Grid;
using Obout.ComboBox;
using System.Collections;

namespace NTSoftReceipt.danhmuc
{
    public partial class dmhuyentinh : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadGrid();
                LoadCombobox();
            }
            AjaxPro.Utility.RegisterTypeForAjax(typeof(NTSoftReceipt.danhmuc.dmhuyentinh));
        }
        public void LoadGrid()
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("SELECT maTinhpr, tenTinh,maHuyenpr,tenHuyen FROM tblDMTinh t LEFT JOIN tblDMQuanHuyen q ON t.maTinhpr=q.maTinhpr_sd WHERE (t.ngungSD = 0 OR t.ngungSD IS NULL) and (q.maHuyenpr != NULL or q.maHuyenpr!=0) ORDER BY tenTinh ASC");
            Grid1.DataBind();
        }
        [AjaxPro.AjaxMethod]
        public void LoadGridWithMaTinh(string maTinh)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("SELECT maTinhpr, tenTinh,maHuyenpr,tenHuyen FROM tblDMTinh t LEFT JOIN tblDMQuanHuyen q ON t.maTinhpr=q.maTinhpr_sd WHERE (t.ngungSD = 0 OR t.ngungSD IS NULL) and (q.maHuyenpr != NULL or q.maHuyenpr!=0) and maTinhpr="+maTinh+" ORDER BY tenTinh ASC");
            Grid1.DataBind();
        }
        public void LoadCombobox()
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Combobox1.DataSource = _sqlfun.GetData("SELECT DISTINCT maTinhpr, tenTinh FROM tblDMTinh t LEFT JOIN tblDMQuanHuyen q ON t.maTinhpr=q.maTinhpr_sd WHERE (t.ngungSD = 0 OR t.ngungSD IS NULL) and (q.maHuyenpr != NULL or q.maHuyenpr!=0) ORDER BY tenTinh ASC");
            Combobox1.DataTextField = "tenTinh";
            Combobox1.DataValueField = "maTinhpr";
            Combobox1.DataBind();
        }

        protected void Combobox1_TextChanged(object sender, EventArgs e)
        {
                ScriptManager.RegisterClientScriptBlock(this, this.GetType(), "alertMessage", "alert('" + Combobox1.SelectedValue + "')", true);
                //LoadGridWithMaTinh(Combobox1.SelectedValue);
        }

        protected void Grid1_Rebind(object sender, EventArgs e)
        {
            SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            Grid1.DataSource = _sqlfun.GetData("SELECT maTinhpr, tenTinh,maHuyenpr,tenHuyen FROM tblDMTinh t LEFT JOIN tblDMQuanHuyen q ON t.maTinhpr=q.maTinhpr_sd WHERE (t.ngungSD = 0 OR t.ngungSD IS NULL) and (q.maHuyenpr != NULL or q.maHuyenpr!=0) and maTinhpr=" + _txtHiddenMatinh.Value + " ORDER BY tenTinh ASC");
            Grid1.DataBind();
        }
    }
}