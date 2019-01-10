using System;
using System.Data;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Xml.Linq;
using System.Web.SessionState;
using WEB_DLL;

namespace NTSoftReceipt.Class
{
    public static class ntsLibraryFunctions
    {
        //Hàm lấy giá trị tự tăng sử dụng AjaxPro
        [AjaxPro.AjaxMethod]
        public static decimal _mGetAutoID(string tableName, string columnName)
        {
            return ntsSqlFunctions._mGetAuToID(tableName, columnName, HttpContext.Current.Session[ntsEnumSessionName.ntsConnectionString1].ToString());
        }
        //Hàm lấy ngày tháng năm định dạng lại sử dụng AjaxPro
        [AjaxPro.AjaxMethod]
        public static string _mConvertDateTime(string DateTime)
        {
            try
            {
                return Convert.ToDateTime(DateTime).ToString("dd/MM/yyyy");
            }
            catch
            { return ""; }
        }
        [AjaxPro.AjaxMethod]
        public static bool _mCheckNumber(string _vstrValue)
        {
            try
            {
                double _vValue = Convert.ToDouble(_vstrValue);
                return true;
            }
            catch
            { }
            return false;
        }
        [AjaxPro.AjaxMethod]
        public static bool _mCheckDetermine(string _vstrValue)
        {
            try
            {
                DateTime _vValue = Convert.ToDateTime(_vstrValue);
                return true;
            }
            catch
            { }
            return false;
        }
    }
}
