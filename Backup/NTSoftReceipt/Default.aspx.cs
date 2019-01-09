using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using NTSoftReceipt.Class;
using NTSoftReceipt.DataConnect;

namespace NTSoftReceipt
{
    public partial class _Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //Tạo các thông số cố định để chạy
            Session.SetConnectionString2(@"Data Source=.\SQLEXPRESS;Initial Catalog=NTSoftReceipt2016;Integrated Security=True");
            Session.SetConnectionString1(@"Data Source=.\SQLEXPRESS;Initial Catalog=UserNTSoftReceipt;Integrated Security=True");
            SqlFunction sqlFun = new SqlFunction(Session.GetConnectionString1());
            UsersDataContext db = new UsersDataContext();
            SqlFunction _vSql = new SqlFunction(Session.GetConnectionString2());
            Session.setNamSudung("2016");
            IQueryable<tblDMDonvi> tblDMDonVi = from tdbDvi in db.tblDMDonvis
                                                where tdbDvi.maDonVi.ToLower() == "89000300"
                                                select tdbDvi;
            tblDMDonvi _vdbDonVi = tblDMDonVi.FirstOrDefault();
            Session.SetDonVi(_vdbDonVi);
            Session.SetDonViCapTren("Đơn vị cấp trên");
            tblUser _vuser = new tblUser();
            _vuser.ngayThaotac = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            _vuser.nguoiThaoTac = 1;
            _vuser.maNguoidungpr = 1;
            _vuser.tenDangNhap = "Admin";
            Session.SetCurrentUser(_vuser);
            Session.SetCurrentDatetime(new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day));
            Session.SetDiaDanh("Vĩnh long");
            Session.SetTenDonVi("Tên đơn vị báo cáo ABC");
            Session.SetHienNgayInBC(true);
            Session.SetNgayInBC("31/12/2016");
        }
    }
}
