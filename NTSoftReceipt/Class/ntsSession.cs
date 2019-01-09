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
using NTSoftReceipt.DataConnect;

namespace NTSoftReceipt.Class
{
    public static class ntsSession
    {
        public static string ntsNamSuDung = "ntsNamSuDung";
        public static string ntsNgayDauKy = "ntsNgayDauKy";
        public static string ntsNgayCuoiKy = "ntsNgayCuoiKy";
        public static string ntsNgayInBC = "ntsNgayInBC";
        public static string ntsHienNgayInBC = "ntsHienNgayInBC";
        public static string ntsMaQHNS = "ntsMaQHNS";

        public static int GetHeight(this HttpSessionState session)
        {
            return (int)session["heightSr"];
        }

        public static void SetHeight(this HttpSessionState session, int value)
        {
            session["heightSr"] = value;
        }

        public static int GetWidth(this HttpSessionState session)
        {
            return (int)session["widthSr"];
        }

        public static void SetWidth(this HttpSessionState session, int value)
        {
            session["widthSr"] = value;
        }

        public static void SetFullScreen(this HttpSessionState session, bool value)
        {
            session[ntsEnumSessionName.ntsFullScreen] = value;
        }

        public static bool GetFullScreen(this HttpSessionState session)
        {
            return (bool)session[ntsEnumSessionName.ntsFullScreen];
        }

        // Gán chuỗi kết nối vào SessionName
        public static void SetConnectionString1(this HttpSessionState session, string sqlConnectionString)
        {
            session[ntsEnumSessionName.ntsConnectionString1] = sqlConnectionString;
        }
        // Lấy chuỗi kết nối từ SessionName
        public static string GetConnectionString1(this HttpSessionState session)
        {
            return session[ntsEnumSessionName.ntsConnectionString1] as String;
        }
        // Thay đổi cơ sở dữ liệu kết nối cho chuỗi kết nối trong SessionName
        public static void ChangeConnectionString1(this HttpSessionState session, string sqlConnectionString, string dbNameSource, string dbNameDes)
        {
            session[ntsEnumSessionName.ntsConnectionString1] = sqlConnectionString.Replace(dbNameSource, dbNameDes);
        }
        // Gán chuỗi kết nối vào SessionName
        public static void SetConnectionString2(this HttpSessionState session, string sqlConnectionString)
        {
            session[ntsEnumSessionName.ntsConnectionString2] = sqlConnectionString;
        }
        // Lấy chuỗi kết nối từ SessionName
        public static string GetConnectionString2(this HttpSessionState session)
        {
            return session[ntsEnumSessionName.ntsConnectionString2] as String;
        }
        // Thay đổi cơ sở dữ liệu kết nối cho chuỗi kết nối trong SessionName
        public static void ChangeConnectionString2(this HttpSessionState session, string sqlConnectionString, string dbNameSource, string dbNameDes)
        {
            session[ntsEnumSessionName.ntsConnectionString2] = sqlConnectionString.Replace(dbNameSource, dbNameDes);
        }
        // Gán người dùng hiện hành vào SessionName
        public static void SetCurrentUser(this HttpSessionState session, tblUser user)
        {
            session[ntsEnumSessionName.ntsCurrentUser] = user;
        }
        /// Lấy người dùng hiện hành từ SessionName
        public static tblUser GetCurrentUser(this HttpSessionState session)
        {
            return session[ntsEnumSessionName.ntsCurrentUser] as tblUser;
        }
        // Lấy mã người sử dụng
        public static decimal GetCurrentUserID(this HttpSessionState session)
        {
            return (session[ntsEnumSessionName.ntsCurrentUser] as tblUser).maNguoidungpr;
        }
        // Gán ngày sử dụng vào SessionName
        public static void SetCurrentDatetime(this HttpSessionState session, DateTime value)
        {
            session[ntsEnumSessionName.ntsCurrentDatetime] = value;
        }
        /// Lấy ngày sử dụng từ SessionName
        public static DateTime GetCurrentDatetime(this HttpSessionState session)
        {
            return Convert.ToDateTime(session[ntsEnumSessionName.ntsCurrentDatetime]);
        }
        // Lấy ngày sử dụng từ SessionName dạng chuỗi có định dạng
        public static string GetCurrentDatetimeMMddyyyy(this HttpSessionState session, string format)
        {
            return Convert.ToDateTime(session[ntsEnumSessionName.ntsCurrentDatetime]).ToString(format);
        }        
        /// Lấy ngày sử dụng từ SessionName dạng chuỗi có định dạng
        public static string GetCurrentDatetimeMMddyyyy(this HttpSessionState session)
        {
            return Convert.ToDateTime(session[ntsEnumSessionName.ntsCurrentDatetime]).ToString("MM/dd/yyyy");
        }
        //Gán Mã đơn vị
        public static void SetDonViCapTren(this HttpSessionState session, string value)
        {
            session["donvicaptren"] = value;
        }
        //Lấy Mã đơn vị
        public static string GetDonViCapTren(this HttpSessionState session)
        {
            return session["donvicaptren"] as string;
        }

        //Gán Mã đơn vị
        public static void SetDonVi(this HttpSessionState session, tblDMDonvi value)
        {
            session["donvi"] = value;
        }
        //Lấy Mã đơn vị
        public static tblDMDonvi GetDonVi(this HttpSessionState session)
        {
            return session["donvi"] as tblDMDonvi;
        }

        public static void SetTenDonVi(this HttpSessionState session, string value)
        {
            session[ntsEnumSessionName.ntsTenDonvi] = value;
        }
        public static string GetTenDonVi(this HttpSessionState session)
        {
            return session[ntsEnumSessionName.ntsTenDonvi] as string;
        }

        public static void SetDiaDanh(this HttpSessionState session, string value)
        {
            session[ntsEnumSessionName.ntsDiadanh] = value;
        }
        public static string GetDiaDanh(this HttpSessionState session)
        {
            return session[ntsEnumSessionName.ntsDiadanh] as string;
        }
        //Lấy năm sử dụng
        public static void setNamSudung(this HttpSessionState session, string value)
        {
            session[ntsNamSuDung] = value;
        }

        public static void SetMaChuong(this HttpSessionState session, string value)
        {
            session[ntsEnumSessionName.ntsMaChuong] = value;
        }
        public static string GetMaChuong(this HttpSessionState session)
        {
            return session[ntsEnumSessionName.ntsMaChuong] as string;
        }

        public static void SetMaQHNS(this HttpSessionState session, string value)
        {
            session[ntsMaQHNS] = value;
        }
        public static string GetMaQHNS(this HttpSessionState session)
        {
            return session[ntsMaQHNS] as string;
        }

        //Lấy năm sử dụng
        public static string GetNamSudung(this HttpSessionState session)
        {
            return session[ntsNamSuDung].ToString();
        }

        //Gán ngay dau kỳ
        public static void SetNgayDauKy(this HttpSessionState session, string value)
        {
            session[ntsNgayDauKy] = value;
        }

        public static string GetNgayDauKy(this HttpSessionState session)
        {
            return session[ntsNgayDauKy].ToString().Substring(3, 2) + "/" + session[ntsNgayDauKy].ToString().Substring(0, 2) + "/" + session[ntsNgayDauKy].ToString().Substring(6, 4);
        }

        public static void SetNgayCuoiKy(this HttpSessionState session, string value)
        {
            session[ntsNgayCuoiKy] = value;
        }

        public static string GetNgayCuoiKy(this HttpSessionState session)
        {
            return session[ntsNgayCuoiKy].ToString().Substring(3, 2) + "/" + session[ntsNgayCuoiKy].ToString().Substring(0, 2) + "/" + session[ntsNgayCuoiKy].ToString().Substring(6, 4);
        }

        public static void SetNgayInBC(this HttpSessionState session, string value)
        {
            session[ntsNgayInBC] = value;
        }

        public static string GetNgayInBC(this HttpSessionState session)
        {
            return session[ntsNgayInBC].ToString().Substring(0, 2) + "/" + session[ntsNgayInBC].ToString().Substring(3, 2) + "/" + session[ntsNgayInBC].ToString().Substring(6, 4);
        }

        public static void SetHienNgayInBC(this HttpSessionState session, bool value)
        {
            session[ntsHienNgayInBC] = value;
        }
        public static bool GetHienNgayInBC(this HttpSessionState session)
        {
            return Convert.ToBoolean(session[ntsHienNgayInBC].ToString());
        }
    }
}
