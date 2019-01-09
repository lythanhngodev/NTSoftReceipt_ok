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
using QLNS2014.Class;
using WEB_DLL;
using System.Data.SqlClient;
using AjaxControlToolkit;
using System.IO;
using Obout.ComboBox;
using Ionic.Zip;
using System.Text.RegularExpressions;
using Obout.Grid;
using System.Data.OleDb;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;

namespace QLNS2014.ThiDuaKhenThuong
{
    public partial class nhapktcanhan : System.Web.UI.Page
    {
        [AjaxPro.AjaxMethod]
        public DataTable dsDanhHieuKT(string laThiDua)
        {
            DataTable tab = new DataTable();
            try
            {
                SqlFunction sqlFunc = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string maDV = HttpContext.Current.Session.GetMaDonVi();
                string strWhereDV3Cap = "maDonVipr_sd  in (" + ntsLibraryFunctions.layDanhSachDonVi_Like(maDV) + @")";
                string sql = @"select '' maKhenThuongpr,N'Tất cả' tenKhenThuong union all  select maKhenThuongpr,tenKhenThuong from tblDMKhenThuong where laThiDua='" + laThiDua + "' and maKhenThuongpr in(select maKhenThuongpr_sd from tblQuyetDinhKT where ( ngayKy between '" + HttpContext.Current.Session.GetNgayDauKy() + "' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + "') and  loaiDoiTuong=N'Cá nhân' and  " + strWhereDV3Cap + ")";
                tab = sqlFunc.GetData(sql);
                return tab;

            }
            catch
            {
                return null;
            }
        }
        [AjaxPro.AjaxMethod]
        public DataTable dsDanhHieuKTAll(string laThiDua)
        {
            SqlFunction sqlFunc = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string maDV = HttpContext.Current.Session.GetMaDonVi();
            string sql = @"select maKhenThuongpr,tenKhenThuong from tblDMKhenThuong where laThiDua='" + laThiDua + "' and loaiDoiTuong LIKE N'%Cá nhân%'";
            return sqlFunc.GetData(sql);
        }
        [AjaxPro.AjaxMethod]
        public string PhanQuyenChucnang()
        {
            return Session["CurrentPermiss"].ToString();
        }
        [AjaxPro.AjaxMethod]
        public bool KiemTraKhoaChucNang()
        {
            return ntsLibraryFunctions.KiemTraKhoaChucNang();
        }
        [AjaxPro.AjaxMethod]
        public DataTable dsHinhThucKTAll(string laThiDua)
        {
            SqlFunction sqlFunc = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string maDV = HttpContext.Current.Session.GetMaDonVi();
            string sql = @"select maHinhThucTDKTpr,tenHinhThucTDKT from tblDMHinhThucTDKT
                            where maHinhThucTDKTpr like '" + laThiDua + "%' and len(isnull(maHinhThucTDKTpr_cha,''))>0";
            return sqlFunc.GetData(sql);
        }
        [AjaxPro.AjaxMethod]
        public string layNgayHT()
        {
            return DateTime.Now.ToString("dd/MM/yyyy");
        }
        [AjaxPro.AjaxMethod]
        public bool kiemTraNgay(string ngay)
        {
            try
            {
                DateTime dt = DateTime.Parse(ngay.ToString().Trim().Split(' ')[0], System.Globalization.CultureInfo.GetCultureInfo("en-gb"));
                if (dt.Year < 1900)
                    return false;
                else
                    return true;
            }
            catch
            {
                return false;
            }
        }
        [AjaxPro.AjaxMethod]
        public string ktraMaDoiTuong(string maDoiTuong, string maDonVi)
        {
            string sql = "select count(sttDoiTuongTDKTpr) from tblDMDoiTuongTDKT where maDoiTuong=N'" + maDoiTuong + "' and maDonVipr_noicongtac=N'" + maDonVi + "'";
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            return sqlFun.GetData(sql).Rows[0][0].ToString();
        }
        [AjaxPro.AjaxMethod]
        public string ktraHoTenVaNamSinh(string tenDoiTuong, string namSinh, string maDonVi)
        {
            if (namSinh == "")
            {
                string sql = "SELECT tenDoiTuongTDKT, ngaySinh FROM tblDMDoiTuongTDKT where tenDoiTuongTDKT=N'" + tenDoiTuong + "' and ngaySinh IS NULL and maDonVipr_noicongtac=N'" + maDonVi + "'";
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                //return sqlFun.GetData(sql).Rows[0][0].ToString();
                if (sqlFun.CheckHasRecord(sql))
                {
                    return "0";
                }
                else return "1";
            }
            else
            {
                if (namSinh.ToString().Length < 5)
                {
                    namSinh = "01/01/" + namSinh.ToString();
                }
                string sql = "SELECT tenDoiTuongTDKT, ngaySinh FROM tblDMDoiTuongTDKT where tenDoiTuongTDKT=N'" + tenDoiTuong + "' and ngaySinh = '" + _mChuyenChuoiSangNgay(namSinh).ToString() + "' and maDonVipr_noicongtac=N'" + maDonVi + "'";
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                //return sqlFun.GetData(sql).Rows[0][0].ToString();
                if (sqlFun.CheckHasRecord(sql))
                {
                    return "0";
                }
                else return "1";
            }
        }
        [AjaxPro.AjaxMethod]
        public bool kiemTraNgayNam(string ngay)
        {
            return kiemtrangaythangnam(ngay);
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            string _pageName = Path.GetFileNameWithoutExtension(Page.AppRelativeVirtualPath);
            Response.AppendHeader("X-XSS-Protection", "0");
            ((CheckBoxColumn)Grid4.Columns[7]).ControlType = GridControlType.Standard;
            AjaxPro.Utility.RegisterTypeForAjax(typeof(QLNS2014.ThiDuaKhenThuong.nhapktcanhan));
            SqlFunction sqlFunc = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            if (!IsPostBack)
            {
                int namSD = Convert.ToInt32(HttpContext.Current.Session.getNamSudung());
                DataTable dtTemp = new DataTable();
                dtTemp.Columns.Add("ten");
                dtTemp.Columns.Add("ma");
                for (int i = namSD - 10; i < namSD + 10; i++)
                {
                    dtTemp.Rows.Add(i.ToString(), i.ToString());
                }
                cboNamXetDuyet.DataSource = dtTemp;
                cboNamXetDuyet.DataTextField = "ten";
                cboNamXetDuyet.DataValueField = "ma";
                cboNamXetDuyet.DataBind();

                string maDV = HttpContext.Current.Session.GetMaDonVi();
                string sql = "select maDonVipr,tenDonVi from tbldmDonVi where  maDonVipr  in (" + ntsLibraryFunctions.layDanhSachDonVi_Like(maDV) + @")";

                cboDonVi_WD.DataSource = sqlFunc.GetData(sql);
                cboDonVi_WD.DataTextField = "tenDonVi";
                cboDonVi_WD.DataValueField = "maDonVipr";
                cboDonVi_WD.DataBind();
                cboDonVi_WD.SelectedIndex = 0;

                cboHinhThucWD1.DataSource = sqlFunc.GetData("SELECT '' maHinhThucTDKTpr,'' tenHinhThucTDKT,'' AS nhomHinhThuc UNION ALL SELECT maHinhThucTDKTpr,tenHinhThucTDKT,nhomHinhThuc=(SELECT tenHinhThucTDKT FROM dbo.tblDMHinhThucTDKT WHERE maHinhThucTDKTpr=a.maHinhThucTDKTpr_cha) FROM dbo.tblDMHinhThucTDKT a WHERE LEN(ISNULL(a.maHinhThucTDKTpr_cha,''))>0  ORDER BY tenHinhThucTDKT");
                cboHinhThucWD1.DataTextField = "tenHinhThucTDKT";
                cboHinhThucWD1.DataValueField = "maHinhThucTDKTpr";
                cboHinhThucWD1.DataBind();
                cboHinhThucWD1.SelectedIndex = 0;


                DataTable tab = new DataTable();
                DataRow dataR;
                tab.Columns.Add(new DataColumn("sttCap", typeof(string)));
                tab.Columns.Add(new DataColumn("tenCap", typeof(string)));
                dataR = tab.NewRow();
                dataR[0] = "1"; dataR[1] = "Cấp Nhà nước";
                tab.Rows.Add(dataR);
                dataR = tab.NewRow();
                dataR[0] = "2"; dataR[1] = "Cấp Bộ, Ngành, Tỉnh";
                tab.Rows.Add(dataR);
                dataR = tab.NewRow();
                dataR[0] = "3"; dataR[1] = "Cấp cơ sở";
                tab.Rows.Add(dataR);
                dataR = tab.NewRow();
                dataR[0] = "4"; dataR[1] = "Cấp đơn vị";
                tab.Rows.Add(dataR);
                dataR = tab.NewRow();
                cboCapQuyetDinh.DataSource = tab;
                cboCapQuyetDinh.DataTextField = "tenCap";
                cboCapQuyetDinh.DataValueField = "sttCap";
                cboCapQuyetDinh.DataBind();

                Grid1.DataSource = null;
                Grid1.DataBind();
                Grid3.DataSource = null;
                Grid3.DataBind();
                Grid4.DataSource = null;
                Grid4.DataBind();
                Grid6.DataSource = null;
                Grid6.DataBind();
                SqlFunction _sqlfun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                DataTable _dtNienDo = _sqlfun.GetData("select nienDoBatDau,nienDoKetThuc from tblDMCauHinhHeThong");
                if (_dtNienDo.Rows.Count > 0)
                {
                    try
                    {
                        DataTable _dtMoi = new DataTable();
                        _dtMoi.Columns.Add("nienDo");
                        _dtMoi.Rows.Add("");
                        for (int i = Convert.ToInt32(_dtNienDo.Rows[0]["nienDoBatDau"].ToString()); i <= Convert.ToInt32(_dtNienDo.Rows[0]["nienDoKetThuc"].ToString()); i++)
                        {
                            _dtMoi.Rows.Add(i);
                        }
                        cboNamXetKhenThuong.DataSource = _dtMoi;
                        cboNamXetKhenThuong.DataTextField = "nienDo";
                        cboNamXetKhenThuong.DataValueField = "nienDo";
                        cboNamXetKhenThuong.DataBind();
                        cboNamXetKhenThuong.SelectedValue = (DateTime.Now.Year.ToString());
                    }
                    catch { }
                }
                cboDanhXung.DataSource = _sqlfun.GetData(@"SELECT '' danhXungpr,'' laTapThe UNION ALL SELECT danhXungpr, laTapThe FROM dbo.tblDMDanhxung WHERE laTapThe = 0");
                cboDanhXung.DataTextField = "danhXungpr";
                cboDanhXung.DataValueField = "danhXungpr";
                cboDanhXung.DataBind();
                //cboPhongBan.DataSource = _sqlfun.GetData(@"SELECT '' sttDoiTuongTDKTpr,'' tenDoiTuongTDKT UNION ALL SELECT CONVERT(NVARCHAR,sttDoiTuongTDKTpr),tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT WHERE loaiDoiTuong=N'Tập thể' and maDonVipr_noicongtac=N'" + cbdonViCongTac.SelectedValue.ToString() + "'");
                //cboPhongBan.DataTextField = "tenDoiTuongTDKT";
                //cboPhongBan.DataValueField = "sttDoiTuongTDKTpr";
                //cboPhongBan.DataBind();

                sql = "SELECT '' maDonVipr,'' maDonViCongTac,'' tenDonViCongTac  UNION ALL select maDonVipr,maDonVipr as maDonViCongTac,tenDonVi as tenDonViCongTac from tbldmDonVi where maDonvipr in (" + ntsLibraryFunctions.layDanhSachDonVi_Like(HttpContext.Current.Session.GetMaDonVi()) + ") ";
                cbdonViCongTac.DataSource = sqlFunc.GetData(sql);
                cbdonViCongTac.DataTextField = "tenDonViCongTac";
                cbdonViCongTac.DataValueField = "maDonVipr";
                cbdonViCongTac.DataBind();

                sql = "SELECT maDonVipr_sd='', tenDonVi=N'Tất cả' UNION ALL SELECT maDonVipr_sd, tenDonVi = (SELECT tenDonVi FROM dbo.tblDMDonvi WHERE maDonvipr = maDonVipr_sd) FROM dbo.tblThiDuaKhenThuong WHERE loaiDoiTuong =N'Cá nhân' and maDonVipr_sd in (" + ntsLibraryFunctions.layDanhSachDonVi_Like(HttpContext.Current.Session.GetMaDonVi()) + ")  GROUP BY maDonVipr_sd";
                cboDonVi.DataSource = sqlFunc.GetData(sql);
                cboDonVi.DataTextField = "tenDonVi";
                cboDonVi.DataValueField = "maDonVipr_sd";
                cboDonVi.DataBind();
                cboDonVi.SelectedIndex = 0;

                cboLinhVuc.DataSource = sqlFunc.GetData("SELECT '' AS maLinhVucpr,'' AS tenLinhVuc UNION ALL SELECT maLinhVucpr,tenLinhVuc FROM dbo.tblDMLinhVuc WHERE ngungSD='0'");
                cboLinhVuc.DataTextField = "tenLinhVuc";
                cboLinhVuc.DataValueField = "maLinhVucpr";
                cboLinhVuc.DataBind();
            }

            AsyncFileUpload.UploadedComplete += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload_UploadedComplete);
            AsyncFileUpload.UploadedFileError += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload_UploadedFileError);
            AsyncFileUpload1.UploadedComplete += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload1_UploadedComplete);
            AsyncFileUpload1.UploadedFileError += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload1_UploadedFileError);
            AsyncFileUpload2.UploadedComplete += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload2_UploadedComplete);
            AsyncFileUpload2.UploadedFileError += new EventHandler<AsyncFileUploadEventArgs>(AsyncFileUpload2_UploadedFileError);
        }
        protected void sdsGioiTinh_Load(object sender, EventArgs e)
        {
            sdsGioiTinh.ConnectionString = Session.GetConnectionString2();
        }
        [AjaxPro.AjaxMethod]
        public bool ktNgayBC(string ngay)
        {
            try
            {
                DateTime.Parse(ngay.ToString().Trim().Split(' ')[0], System.Globalization.CultureInfo.GetCultureInfo("en-gb"));
                return true;
            }
            catch
            {
                return false;
            }
        }
        [AjaxPro.AjaxMethod]
        public string taoMaDoiTuong1(string maDonVi)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string soPhieu = "";
            string sql = "SELECT MAX(CONVERT(DECIMAL,RIGHT(maDoiTuong,6))) FROM dbo.tblDMDoiTuongTDKT WHERE maDoiTuong like N'" + maDonVi + "-%' and ISNUMERIC(RIGHT(maDoiTuong,6)) = 1 AND maDonvipr_noicongtac=N'" + maDonVi + "'";
            decimal _vNewKey = sqlFun.GetOneDecimalField(sql) + 1;
            soPhieu = maDonVi + "-" + _vNewKey.ToString("000000");
            return soPhieu;
        }
        [AjaxPro.AjaxMethod]
        public string kiemTraThongTinQD(string soQD, string ngayQD, string capQD)
        {
            SqlFunction sqlFunc = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string sql = "select soQuyetDinh from tblThiDuaKhenThuong where soQuyetDinh=N'" + soQD + "' and ngayKyQD='" + _mChuyenChuoiSangNgay(ngayQD) + "' and capQuyetDinh=N'" + capQD + "' and loaiDoiTuong=N'Cá nhân' and ngayDangKy between '" + HttpContext.Current.Session.GetNgayDauKy() + "' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + "'";
            return sqlFunc.GetOneStringField(sql);
        }

        [AjaxPro.AjaxMethod]
        public string getFile(string tenFile)
        {
            string hostName = HttpContext.Current.Request.Url.Host;
            string urlFileHtml = "http://view.officeapps.live.com/op/view.aspx?src=http://" + hostName + "/ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + "/" + tenFile;
            return urlFileHtml;
        }
        [AjaxPro.AjaxMethod]
        public string getError()
        {
            if (Session["ktrDL"] == null)
            {
                return "";
            }
            else
            {
                return Session["ktrDL"].ToString();
            }
        }
        [AjaxPro.AjaxMethod]
        public string taiFile(string fileName)
        {
            try
            {
                string hostName = HttpContext.Current.Request.Url.Host;
                string urlFileHtml = "/ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + "/" + fileName;
                return urlFileHtml;
            }
            catch
            {
                return "";
            }
        }
        [AjaxPro.AjaxMethod]
        public void xoaQuyetDinh(string sttQuyetDinhKTpr, string filename)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            sqlFun.ExeCuteNonQuery("Update tblQuyetDinhKT set fileDinhKem='' where sttQuyetDinhKTpr='" + sttQuyetDinhKTpr + "'");
            System.IO.File.Delete(string.Concat(Server.MapPath("~/ThiDuaKhenThuong/HinhAnhTDKT/" + HttpContext.Current.Session.getNamSudung() + "/" + filename)));
        }
        [AjaxPro.AjaxMethod]
        public void xoaFileDinhKem(string duongDan)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            sqlFun.ExeCuteNonQuery("delete tblHinhAnhKemTheo where duongDan=N'" + duongDan + "'");
            System.IO.File.Delete(string.Concat(Server.MapPath(duongDan)));
        }
        public void Grid1_OnRebind(object sender, EventArgs e)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());

                string strWhereDanhHieu = "";
                if (cboSearchDanhHieu.SelectedValue != "")
                    strWhereDanhHieu = " maKhenThuongpr_sd=N'" + cboSearchDanhHieu.SelectedValue + "' and ";
                string s = @"SELECT [sttQuyetDinhKTpr] , ngayKhenThuong =convert(varchar(10),ngayKhenThuong,103)
                      ,capQD=(case when capQuyetDinh=1 then N'Cấp Nhà nước'
				           when capQuyetDinh=2 then N'Cấp Bộ, Ngành, Tỉnh'
				           when capQuyetDinh=3 then N'Cấp cơ sở'
				           when capQuyetDinh=4 then N'Cấp đơn vị' end)
                      ,tenKhenThuong=(select tenKhenThuong from tblDMKhenThuong where maKhenThuongpr=maKhenThuongpr_sd)
                      ,[maKhenThuongpr_sd],fileDinhKem,sttHoiDongXDTDpr_sd,namXetKhenThuong
                      ,fileDinhKem1= case when len(isnull(fileDinhKem,''))>0 then 
						'<a href =''javascript:void(0)'' onclick=''taiQuyetDinh('+Char(34) + fileDinhKem + Char(34)+N');return false;''>Tải quyết định</a>'
						else '' end
                      ,hinhThucKhenThuong=(select tenhinhThucTDKT from tblDMHinhThucTDKT where maHinhThucTDKTpr=maHinhThucTDKTpr_sd)
                      ,[maHinhThucTDKTpr_sd] ,[moTa] ,[soQuyetDinh] , ngayKy=convert(varchar(10),ngayKy,103) ,[capQuyetDinh] ,[nguoiKy] ,[trichYeu] ,[noiDung], maLinhVucpr_sd, tenLinhVucKT=(SELECT tenLinhVuc FROM dbo.tblDMLinhVuc WHERE maLinhVucpr=tblQuyetDinhKT.maLinhVucpr_sd)
                       FROM tblQuyetDinhKT where ISNULL(danhGiaCongChuc,0) <> 1 and loaiDoiTuong=N'Cá nhân' and  ( ngayKy between '" + HttpContext.Current.Session.GetNgayDauKy() + "' and '" + HttpContext.Current.Session.GetNgayCuoiKy() +
                       "') and" + (cboSearchDanhHieu.SelectedValue == "" ? "" : " maKhenThuongpr_sd=N'" + cboSearchDanhHieu.SelectedValue + "' and ") + @" " + (cboDonVi.SelectedValue.ToString() == "" ? "" : " maDonVipr_sd=N'" + cboDonVi.SelectedValue + "' and ") + @"
                       maKhenThuongpr_sd in (select maKhenThuongpr from tblDMKhenThuong where " + (rdKhenThuong.Checked ? "laThiDua=0" : "laThiDua=1") + ")";
                Grid1.DataSource = sqlFun.GetData(s);
                Grid1.DataBind();
            }
            catch
            {
                Grid1.DataSource = null;
                Grid1.DataBind();
            }
        }
        public void Grid4_OnRebind(object sender, EventArgs e)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string maDV = HttpContext.Current.Session.GetMaDonVi();// cbdonViCongTac.SelectedValue;
                //string strWhereDV = " maDonVipr_sd IN (SELECT maDonvipr FROM dbo.tblDMDonvi WHERE (maDonvipr= N'" + maDV + "' OR (maDonvipr_cha= N'" + maDV + "') OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "')  OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "'))) OR  maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "') ) ) )";
                //string strWhereDV = " sttDoiTuongTDKTpr_sd IN(SELECT sttDoiTuongTDKTpr FROM dbo.tblDMDoiTuongTDKT WHERE maDonVipr_noicongtac=N'" + maDV + "')";
                string strWhereDanhHieu = "";
                if (cboSearchDanhHieu.SelectedValue != "")
                    strWhereDanhHieu = " maKhenThuongpr_sd=N'" + cboSearchDanhHieu.SelectedValue + "' and ";
                string s = @" SELECT ROW_NUMBER()OVER(ORDER BY doiTuong) AS stt,'' AS temp1,'' AS temp2,* FROM(
                                    SELECT case when CHARINDEX(' ',doiTuong,1) > 0 then ltrim(right(doiTuong,CHARINDEX(' ',REVERSE(doiTuong),1))) else doiTuong end as serachName,* FROM (
                                    select *,
                                    case when danhHieu is null then hinhThuc else case when hinhThuc is null then danhHieu else danhHieu +'/'+ hinhThuc end end as danhhieuhinhthuc
                                    from (
                                    select ngayDangKy=convert(varchar(10),ngayDangKy,103),ngayKhenThuong=convert(varchar(18),ngayKhenThuong,103),ngayKyQD=convert(varchar(18),ngayKyQD,103),maHinhThucTDKTpr_sd,maKhenThuongpr_sd,sttTDKTpr,capQuyetDinh,soQuyetDinh,nguoiKy,noiDungTDKT,trichYeu,moTa,kemTheoQD,
                                    (select tenDoiTuongTDKT from dbo.tblDMDoiTuongTDKT where sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd) as doiTuong,
                                    (select tenChucVuQL from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd) as chucVu,
                                    (select ngaySinh = (case when namSinh ='1'  then convert(nvarchar(4),year(ngaySinh)) else convert(nvarchar(10),ngaysinh,103) end) from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd) as ngaySinh ,
                                    noiCongTac=(SELECT (SELECT tenDonVi FROM dbo.tblDMDonvi WHERE maDonvipr=a.maDonVipr_noicongtac) FROM dbo.tblDMDoiTuongTDKT a WHERE sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd),
                                    (select ngayThanhLap from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd) as ngayThanhLap ,
                                    (select kyHieu from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd) as kyHieu,
                                    (select tenKhenThuong from dbo.tblDMKhenThuong where maKhenThuongpr = maKhenThuongpr_sd) as danhHieu,
                                    (select tenHinhThucTDKT from dbo.tblDMHinhThucTDKT where maHinhThucTDKTpr = maHinhThucTDKTpr_sd) as hinhThuc,
                                    taiFile='<a href =''javascript:void(0)'' onclick=''xemFile('+Char(34) + kemTheoQD + Char(34)+');return false;''>Xem QĐ</a>'
                                    ,hinhAnhTD=(N'<a href =''javascript:void(0)'' onclick=''moWindow('+ CONVERT(varchar(18), sttTDKTpr)+N');return false''>Xem ảnh</a>')
                                    ,tenCapKT=
                                                (
	                                                case when capQuyetDinh=1 then N'Cấp Nhà nước'
		                                                 when capQuyetDinh=2 then N'Cấp Bộ, Ngành, Tỉnh'
		                                                 when capQuyetDinh=3 then N'Cấp cơ sở'
		                                                 when capQuyetDinh=4 then N'Cấp đơn vị'
		                                                 end
                                                )
,checkDauKy= case when sttDKThiDuapr_sd is NULL then 
						1
						else 0 end
                                    ,xemHinh='<a href =''javascript:void(0)'' onclick=''xemAnh('+Char(34) + convert(varchar(18),sttTDKTpr) + Char(34)+N');return false;''>Xem ảnh</a>'
                                    from tblThiDuaKhenThuong where sttQuyetDinhKTpr_sd='" + sttQuyetDinhKTpr_sd.Value + @"') as temp  )AS temp2  )AS temp3 WHERE  serachName COLLATE SQL_Latin1_General_CP1_CI_AI LIKE N'" + hdfAlphabet.Value.ToString() + "%'";
                Grid4.DataSource = sqlFun.GetData(s);
                Grid4.DataBind();
            }
            catch
            {

            }
        }

        void AsyncFileUpload1_UploadedComplete(object sender, AsyncFileUploadEventArgs e)
        {
            try
            {
                string path = "";
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string[] arr = e.filename.Split('\\');

                string fileName = arr[arr.Length - 1];// e.filename;

                if (!System.IO.Directory.Exists(Server.MapPath("~/ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + "")))
                {
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + ""));
                }
                if (Path.GetExtension(e.filename).Contains(".pdf") == false)
                {
                    return;
                }
                string strDate = DateTime.Now.ToString("ddmmyyyyhhmmss");
                string fileExtension = Path.GetExtension(fileName).Replace(".", "");
                fileName = fileName.Substring(fileName.LastIndexOf("\\\\") + 1);
                fileName = fileName.Substring(0, fileName.LastIndexOf(fileExtension) - 1) + "_" + strDate + "." + fileExtension;
                fileName = fileName.Replace(" ", "").Replace(";", "");
                path = string.Concat(Server.MapPath("~/ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + "/" + fileName));
                AsyncFileUpload1.SaveAs(path);
                sqlFun.ExeCuteNonQuery("update tblQuyetDinhKT set fileDinhKem=N'" + fileName + "' where sttQuyetDinhKTpr='" + sttQuyetDinhKTpr_sd.Value + "'");
                Session.Add("filePdfUpload", fileName);
            }
            catch
            {
            }
            AsyncFileUpload1.Dispose();
        }


        void AsyncFileUpload1_UploadedFileError(object sender, AsyncFileUploadEventArgs e)
        {
        }
        protected void Grid3_OnRebind(object sender, EventArgs e)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                string sqlQuery = @"select *
                    from(select 
                    serachName=(select case when CHARINDEX(' ',tenDoiTuongTDKT,1) > 0 
                    then ltrim(right(tenDoiTuongTDKT,CHARINDEX(' ',REVERSE(tenDoiTuongTDKT),1))) else tenDoiTuongTDKT end),*
                    from(select sttTDKTpr, sttDoiTuongTDKTpr_sd, sttDoiTuongTDKTpr_cha,maKhenThuongpr_sd,ngayDangKy=convert(varchar(10),ngayDangKy,103)
                    ,maDoiTuong=(select maDoiTuong from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)	
                    ,tenDoiTuongTDKT=(select tenDoiTuongTDKT from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)
                    ,phongBan=(SELECT tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT b WHERE b.sttDoiTuongTDKTpr = isnull(tblThiDuaKhenThuong.sttDoiTuongTDKTpr_cha,0))
                    ,donViCongTac=(select tenDonVi from tblDMDonvi where  maDonvipr=(select maDonVipr_noicongtac from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd))
                    ,maDonVipr_noicongtac = (select maDonVipr_noicongtac from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)
                    ,checkDauKy= case when sttDKThiDuapr_sd is NULL then 
						1
						else 0 end
                    ,xemHinh='<a href =''javascript:void(0)'' onclick=''xemAnh('+Char(34) + convert(varchar(18),sttTDKTpr) + Char(34)+N');return false;''>Xem ảnh</a>'
                    ,maLinhVucpr_sd,thamNienQD,ghiChu
                    from tblThiDuaKhenThuong where sttQuyetDinhKTpr_sd='" + sttQuyetDinhKTpr_sd.Value + "')as temp)as temp1 where serachName COLLATE SQL_Latin1_General_CP1_CI_AI LIKE N'" + hdfAlphabet1.Value.ToString() + "%'";
                Grid3.DataSource = sqlFun.GetData(sqlQuery);
                Grid3.DataBind();
            }
            catch
            {
                Grid3.DataSource = null;
                Grid3.DataBind();
            }
        }
        [AjaxPro.AjaxMethod]
        public DataTable dsHinhAnh(string sttTDKT)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            return sqlFun.GetData("select sttHinhAnhpr,duongDan from tblHinhAnhKemTheo where sttTDKTpr_sd='" + sttTDKT + "'");
        }
        protected void Grid6_OnRebind(object sender, EventArgs e)
        {
            try
            {
                Grid6.DataSource = xuLyTableDoiTuong(ngayKy.Text);
                Grid6.DataBind();
            }
            catch
            {
                Grid6.DataSource = null;
                Grid6.DataBind();
            }
        }
        private string _mChuyenChuoiSangNgay(string ddMMyyyy)
        {
            return ddMMyyyy.Substring(3, 2) + "/" + ddMMyyyy.Substring(0, 2) + "/" + ddMMyyyy.Substring(6, 4);
        }

        [AjaxPro.AjaxMethod]
        public bool xoaQuyetDinhKT(string sttQuyetDinh)
        {
            DataTable tab = new DataTable();
            bool bolTraVe, bolTraVe2 = true, bolTraVe3 = true;
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                string sqlQuery = @"select *
                    from(select 
                    serachName=(select case when CHARINDEX(' ',tenDoiTuongTDKT,1) > 0 
                    then ltrim(right(tenDoiTuongTDKT,CHARINDEX(' ',REVERSE(tenDoiTuongTDKT),1))) else tenDoiTuongTDKT end),*
                    from(select sttTDKTpr, sttDoiTuongTDKTpr_sd, maKhenThuongpr_sd,ngayDangKy=convert(varchar(10),ngayDangKy,103)
                    ,maDoiTuong=(select maDoiTuong from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)	
                    ,tenDoiTuongTDKT=(select tenDoiTuongTDKT from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)
                    ,phongBan=(select tenPhongBan from tblDMPhongBan where sttPhongBanpr=(select sttPhongBanpr_sd from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd))
                    ,donViCongTac=(select tenDonVi from tblDMDonvi where  maDonvipr=(select maDonVipr_noicongtac from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd))
                    ,xemHinh='<a href =''javascript:void(0)'' onclick=''xemAnh('+Char(34) + convert(varchar(18),sttTDKTpr) + Char(34)+N');return false;''>Xem ảnh</a>'
                    from tblThiDuaKhenThuong where sttQuyetDinhKTpr_sd='" + sttQuyetDinh + "')as temp)as temp1";
                tab = sqlFun.GetData(sqlQuery);
                if (tab.Rows.Count > 0)
                {
                    try
                    {
                        for (int i = 0; i < tab.Rows.Count; i++)
                        {
                            bolTraVe = xoaDoiTuongTDKT(tab.Rows[i][1].ToString());
                            if (!bolTraVe)
                                return bolTraVe2 = false;
                        }
                        if (bolTraVe2)
                        {
                            try
                            {
                                SqlConnection sqlCon = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                                sqlCon.Open();
                                SqlCommand cmd = new SqlCommand("delete tblQuyetDinhKT where sttQuyetDinhKTpr=@sttQuyetDinhKTpr", sqlCon);
                                cmd.Parameters.AddWithValue("@sttQuyetDinhKTpr", sttQuyetDinh);
                                cmd.ExecuteNonQuery();
                                sqlCon.Close();
                                bolTraVe3 = true;
                            }
                            catch
                            {
                                bolTraVe3 = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        bolTraVe3 = false;
                    }
                }
                else
                {
                    try
                    {
                        SqlConnection sqlCon = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                        sqlCon.Open();
                        SqlCommand cmd = new SqlCommand("delete tblQuyetDinhKT where sttQuyetDinhKTpr=@sttQuyetDinhKTpr", sqlCon);
                        cmd.Parameters.AddWithValue("@sttQuyetDinhKTpr", sttQuyetDinh);
                        cmd.ExecuteNonQuery();
                        sqlCon.Close();
                        bolTraVe3 = true;
                    }
                    catch
                    {
                        bolTraVe3 = false;
                    }
                }
            }
            catch
            {
                bolTraVe3 = false;
            }

            return bolTraVe3;
        }
        [AjaxPro.AjaxMethod]
        public string kiemTraXoa(string key, string ma)
        {
            return ntsLibraryFunctions.KiemTraXoa(ma, "'tblThiDuaKhenThuong'", key, "sttTDKTpr");
        }
        [AjaxPro.AjaxMethod]
        public bool xoaDoiTuongTDKT(string sttTDKT)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                if (sqlFun.GetOneStringField(@"select tenKhenThuong from tblDMKhenThuong where maKhenThuongpr=
                             (select maKhenThuongpr_cha from tblDMKhenThuong where maKhenThuongpr=
                            (select maKhenThuongpr_sd from tblThiDuaKhenThuong where sttTDKTpr='" + sttTDKT + "'))") == "Khen thưởng")
                {
                    sqlFun.ExeCuteNonQuery("delete tblThiDuaKhenThuong WHERE sttTDKTpr='" + sttTDKT + "'");
                }
                else
                {
                    if (sqlFun.GetOneDecimalField(@"SELECT sttTDKTpr FROM dbo.tblThiDuaKhenThuong where sttTDKTpr='" + sttTDKT + "' AND sttDKThiDuapr_sd IS NOT NULL ") == 0)
                    {
                        sqlFun.ExeCuteNonQuery("delete tblThiDuaKhenThuong WHERE sttTDKTpr='" + sttTDKT + "'");
                    }
                    else
                    {
                        sqlFun.ExeCuteNonQuery("UPDATE tblThiDuaKhenThuong set namXetKhenThuong=null,ngayKhenThuong=null, moTa=null,soQuyetDinh=null,ngayKyQD=null,capQuyetDinh=null,noiDungTDKT=null,tinhTrang=null,kemTheoQD=null,trichYeu=null,nguoiKy=null,sttQuyetDinhKTpr_sd=null WHERE sttTDKTpr='" + sttTDKT + "'");
                    }
                }
                sqlFun.ExeCuteNonQuery("delete tblHinhAnhKemTheo where sttTDKTpr_sd='" + sttTDKT + "'");
                return true;
            }
            catch (Exception EX)
            {
                return false;
            }
        }

        [AjaxPro.AjaxMethod]
        public string luuThongTin(object[] param, string flag)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                SqlConnection sqlCon = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                sqlCon.Open();
                string result = "";
                if (flag == "0")
                    result = @"INSERT INTO [dbo].[tblQuyetDinhKT] ([ngayKhenThuong], [maKhenThuongpr_sd], [maHinhThucTDKTpr_sd], [moTa], [soQuyetDinh], [ngayKy], [capQuyetDinh], [nguoiKy], [trichYeu], [noiDung], [ngayThaoTac], [nguoiThaoTac], [maDonVipr_sd],loaiDoiTuong,sttHoiDongXDTDpr_sd,namXetKhenThuong,maLinhVucpr_sd)
	                    SELECT  @ngayKhenThuong, @maKhenThuongpr_sd, @maHinhThucTDKTpr_sd, @moTa, @soQuyetDinh, @ngayKy, @capQuyetDinh, @nguoiKy, @trichYeu, @noiDung, GETDATE(), @nguoiThaoTac, @maDonVipr_sd,N'Cá nhân',@sttHoiDongXDTDpr_sd,@namXetKhenThuong,@maLinhVucpr_sd";
                else
                    result = @"UPDATE [dbo].[tblQuyetDinhKT]
	                    SET  [ngayKhenThuong] = @ngayKhenThuong, [maKhenThuongpr_sd] = @maKhenThuongpr_sd, [maHinhThucTDKTpr_sd] = @maHinhThucTDKTpr_sd, [moTa] = @moTa, [soQuyetDinh] = @soQuyetDinh, [ngayKy] = @ngayKy, [capQuyetDinh] = @capQuyetDinh, [nguoiKy] = @nguoiKy, [trichYeu] = @trichYeu, [noiDung] = @noiDung,sttHoiDongXDTDpr_sd=@sttHoiDongXDTDpr_sd,namXetKhenThuong=@namXetKhenThuong,maLinhVucpr_sd=@maLinhVucpr_sd
	                    WHERE  [sttQuyetDinhKTpr] = @sttQuyetDinhKTpr; update tblThiDuaKhenThuong set [ngayKhenThuong] = @ngayKhenThuong, [maKhenThuongpr_sd] = @maKhenThuongpr_sd, [maHinhThucTDKTpr_sd] = @maHinhThucTDKTpr_sd, [moTa] = @moTa, [soQuyetDinh] = @soQuyetDinh, [ngayKyQD] = @ngayKy, [capQuyetDinh] = @capQuyetDinh, [nguoiKy] = @nguoiKy, [trichYeu] = @trichYeu, [noiDungTDKT] = @noiDung,namXetKhenThuong=@namXetKhenThuong,maLinhVucpr_sd=@maLinhVucpr_sd
	                    WHERE  [sttQuyetDinhKTpr_sd] = @sttQuyetDinhKTpr";
                SqlCommand cmd = new SqlCommand(result, sqlCon);
                if (flag != "0")
                    cmd.Parameters.Add(new SqlParameter("@sttQuyetDinhKTpr", param[0].ToString()));
                cmd.Parameters.Add(new SqlParameter("@ngayKhenThuong", _mChuyenChuoiSangNgay(param[3].ToString())));
                cmd.Parameters.Add(new SqlParameter("@maKhenThuongpr_sd", param[6].ToString()));
                cmd.Parameters.Add(new SqlParameter("@maHinhThucTDKTpr_sd", param[9].ToString()));
                cmd.Parameters.Add(new SqlParameter("@moTa", param[1].ToString()));
                cmd.Parameters.Add(new SqlParameter("@soQuyetDinh", param[2].ToString()));
                cmd.Parameters.Add(new SqlParameter("@ngayKy", _mChuyenChuoiSangNgay(param[3].ToString())));
                cmd.Parameters.Add(new SqlParameter("@capQuyetDinh", param[4].ToString()));
                cmd.Parameters.Add(new SqlParameter("@nguoiKy", param[8].ToString()));
                cmd.Parameters.Add(new SqlParameter("@trichYeu", param[7].ToString()));
                cmd.Parameters.Add(new SqlParameter("@noiDung", param[5].ToString()));
                cmd.Parameters.Add(new SqlParameter("@sttHoiDongXDTDpr_sd", param[11].ToString()));
                cmd.Parameters.Add(new SqlParameter("@namXetKhenThuong", param[12].ToString()));
                cmd.Parameters.Add(new SqlParameter("@maLinhVucpr_sd", param[13].ToString()));
                cmd.Parameters.Add(new SqlParameter("@nguoiThaoTac", HttpContext.Current.Session.GetCurrentUserID()));
                cmd.Parameters.Add(new SqlParameter("@maDonVipr_sd", HttpContext.Current.Session.GetMaDonVi()));
                if (cmd.ExecuteNonQuery() > 0)
                {
                    return sqlFun.GetOneStringField("select CONVERT(varchar(18),max(sttQuyetDinhKTpr)) from tblQuyetDinhKT where maDonVipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'");
                }

                else
                {
                    return "0";
                }
            }
            catch
            {
                return "0";
            }
        }
        [AjaxPro.AjaxMethod]
        public string chonDoiTuong(object[] param, string flag)
        {

            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            SqlConnection sqlCon = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            sqlCon.Open();
            string sql = "";
            string result = "";
            try
            {


                if ((sqlFun.GetOneStringField(@"select tenKhenThuong from tblDMKhenThuong where maKhenThuongpr= (select maKhenThuongpr_cha from tblDMKhenThuong where maKhenThuongpr='" + param[5].ToString() + "')") == "Khen thưởng") || (bool.Parse(param[13].ToString()) == true))
                {
                    try
                    {
                        string[] separators = { "," };
                        string[] doiTuong = param[12].ToString().Replace("'", "").Split(separators, StringSplitOptions.RemoveEmptyEntries);
                        string fileUpload = "";
                        if (Session["filePdfUpload"] != null)
                            //fileUpload = Session["filePdfUpload"].ToString();
                            fileUpload = "../ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + "/" + Session["filePdfUpload"].ToString();
                        for (int i = 0; i < doiTuong.Length; i++)
                        {
                            string stttapthe = sqlFun.GetOneStringField(@"select convert(nvarchar(50),sttDoiTuongTDKTpr_cha) from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr ='" + doiTuong[i].ToString() + "'");
                            result += @" SELECT " + (stttapthe == "" ? "null" : stttapthe) + ",null, " + doiTuong[i].ToString() + ",null, null, null, N'" + param[5].ToString() + "', @moTa, @soQuyetDinh, @ngayKyQD, @capQuyetDinh, @noiDungTDKT, '" + param[9].ToString() + @"', @kemTheoQD, N'Cá nhân', N'Có quyết định', N'" + HttpContext.Current.Session.GetMaDonVi() + @"' 
                        , getdate(), " + HttpContext.Current.Session.GetCurrentUserID() + ", null, null, null, @trichYeu, @nguoiky, null, null, @ngayKhenThuong,@sttQuyetDinhKTpr_sd,@namXetKhenThuong,N'" + sqlFun.GetOneStringField(@"select maDonVipr_noicongtac from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr ='" + doiTuong[i].ToString() + "'") + "',N'" + sqlFun.GetOneStringField(@"select tenChucVuQL from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr ='" + doiTuong[i].ToString() + "'") + "',@maLinhVucpr_sd union all";
                        }
                        result = "INSERT INTO tblThiDuaKhenThuong(sttDoiTuongTDKTpr_cha,sttDKThiDuapr_sd, sttDoiTuongTDKTpr_sd,ngayDangKy, noiDungDK, kemTheoBCThanhTich, maKhenThuongpr_sd, moTa, soQuyetDinh, ngayKyQD, capQuyetDinh, noiDungTDKT, maHinhThucTDKTpr_sd, kemTheoQD, loaiDoiTuong, tinhTrang, maDonVipr_sd, ngayThaoTac, nguoiThaoTac, sttPhongBanpr_sd, ghiChu, nhapExcel, trichYeu, nguoiky, nhapExcelKQ, thuTuUuTien, ngayKhenThuong,sttQuyetDinhKTpr_sd,namXetKhenThuong,maDonVipr_noicongtac,tenChucVu,maLinhVucpr_sd ) " + result.Substring(0, result.Length - 10);

                    }
                    catch
                    {
                        return "0";
                    }
                }
                else
                {
                    result = @"UPDATE tblThiDuaKhenThuong 
                            set ngayKhenThuong=@ngayKhenThuong, moTa=@moTa,soQuyetDinh=@soQuyetDinh,ngayKyQD=@ngayKyQD
                            ,capQuyetDinh=@capQuyetDinh,noiDungTDKT=@noiDungTDKT,tinhTrang=@tinhTrang,kemTheoQD=@kemTheoQD
                            ,trichYeu=@trichYeu,nguoiKy=@nguoiKy,sttQuyetDinhKTpr_sd=@sttQuyetDinhKTpr_sd,namXetKhenThuong=@namXetKhenThuong,maLinhVucpr_sd=@maLinhVucpr_sd
                            WHERE sttTDKTpr in (" + param[11].ToString() + ")";

                }
                try
                {
                    SqlCommand cmd = new SqlCommand(result, sqlCon);
                    cmd.Parameters.Add(new SqlParameter("@moTa", param[0].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@soQuyetDinh", param[1].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@sttQuyetDinhKTpr_sd", param[6].ToString()));
                    if (string.IsNullOrEmpty(param[2].ToString().Trim()))
                        cmd.Parameters.Add(new SqlParameter("@ngayKyQD", DBNull.Value));
                    else
                        cmd.Parameters.Add(new SqlParameter("@ngayKyQD", _mChuyenChuoiSangNgay(param[2].ToString())));
                    if (string.IsNullOrEmpty(param[10].ToString().Trim()))
                        cmd.Parameters.Add(new SqlParameter("@ngayKhenThuong", DBNull.Value));
                    else
                        cmd.Parameters.Add(new SqlParameter("@ngayKhenThuong", _mChuyenChuoiSangNgay(param[2].ToString())));
                    cmd.Parameters.Add(new SqlParameter("@capQuyetDinh", param[3].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@noiDungTDKT", param[4].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@tinhTrang", "Có quyết định"));
                    if (Session["filePdfUpload"] == null)
                        cmd.Parameters.Add(new SqlParameter("@kemTheoQD", ""));
                    else
                        //cmd.Parameters.Add(new SqlParameter("@kemTheoQD", Session["filePdfUpload"].ToString()));
                        cmd.Parameters.Add(new SqlParameter("@kemTheoQD", "../ThiDuaKhenThuong/quyetDinhKT/" + HttpContext.Current.Session.getNamSudung() + "/" + Session["filePdfUpload"].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@trichYeu", param[7].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@nguoiKy", param[8].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@namXetKhenThuong", param[14].ToString()));
                    cmd.Parameters.Add(new SqlParameter("@maLinhVucpr_sd", param[15].ToString()));
                    cmd.ExecuteNonQuery();
                    return "1";
                }
                catch
                {
                    return "0";
                }
            }
            catch
            {
                return "0";
            }
        }
        [AjaxPro.AjaxMethod]
        public string getPDFFileName()
        {
            try
            {
                return Session["filePdfUpload"].ToString();
            }
            catch
            {
                return "";
            }
        }

        public DataTable xuLyTableDoiTuong(string ngayDangKy)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            DataTable tab = new DataTable();
            string maDV = cboDonVi_WD.SelectedValue;
            if (maDV == "") maDV = HttpContext.Current.Session.GetMaDonVi();
            //Kiểm tra danh hiệu nhập
            string sqlQuery = "";
            //Nếu nhập khen thưởng hoặc nhập đầu kỳ thì load hết danh mục cá nhân
            if (chkDauKy.Checked)
            {
                sqlQuery = @"select *,(SELECT ngaySinh FROM tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_SD) AS ngaySinh,tenKhenThuong='' from(select 
                            serachName=(select case when CHARINDEX(' ',tenDoiTuongTDKT,1) > 0 
                            then ltrim(right(tenDoiTuongTDKT,CHARINDEX(' ',REVERSE(tenDoiTuongTDKT),1))) else tenDoiTuongTDKT end),*
                            from(select sttDoiTuongTDKTpr_sd=sttDoiTuongTDKTpr
                            ,maDoiTuong
                            ,tenDoiTuongTDKT
                            --,phongBan=(select tenDoiTuongTDKT from tblDMDoiTuongTDKT b where b.sttDoiTuongTDKTpr = tblDMDoiTuongTDKT.sttDoiTuongTDKTpr_cha)
                            ,donViCongTac=(select tenDonVi from tblDMDonvi where  maDonvipr=maDonVipr_noicongtac)
                            ,ketQua=N'Đủ điều kiện'
                            from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr NOT IN (SELECT sttDoiTuongTDKTpr_sd FROM dbo.tblThuyenChuyen WHERE ngayThuyenChuyen<'" + _mChuyenChuoiSangNgay(ngayDangKy) + @"' AND maDonVipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + @"')  and
                            sttDoiTuongTDKTpr not in(select sttDoiTuongTDKTpr_sd from tblThiDuaKhenThuong where sttQuyetDinhKTpr_sd='" + sttQuyetDinhKTpr_sd.Value + @"') and
                            loaiDoiTuong=N'Cá nhân' and maDonVipr_noicongtac= N'" + maDV + "')as temp)as temp1 where serachName COLLATE SQL_Latin1_General_CP1_CI_AI LIKE N'" + hdfAlphabet1.Value.ToString() + "%'";
                return sqlFun.GetData(sqlQuery);
            }
            else
            {
                sqlQuery = @"select *,(SELECT ngaySinh FROM tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_SD) AS ngaySinh,tenKhenThuong=(SELECT tenKhenThuong +'; ' FROM tblDMKhenThuong WHERE maKhenThuongpr in  (select s.Data from fnSplit(';',temp1.maKhenThuongpr_sd) s)  FOR XML PATH ('')) from(select 
                    serachName=(select case when CHARINDEX(' ',tenDoiTuongTDKT,1) > 0 
                    then ltrim(right(tenDoiTuongTDKT,CHARINDEX(' ',REVERSE(tenDoiTuongTDKT),1))) else tenDoiTuongTDKT end),*
                    from(select sttTDKTpr, sttDoiTuongTDKTpr_sd, maKhenThuongpr_sd
                    ,maDoiTuong=(select maDoiTuong from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)	
                    ,tenDoiTuongTDKT=(select tenDoiTuongTDKT from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd)
                     --,phongBan=(select tenDoiTuongTDKT from tblDMDoiTuongTDKT b where b.sttDoiTuongTDKTpr = tblDMDoiTuongTDKT.sttDoiTuongTDKTpr_cha)
                    ,donViCongTac=(select tenDonVi from tblDMDonvi where  maDonvipr=(select maDonVipr_noicongtac from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd))
                     from tblThiDuaKhenThuong where sttQuyetDinhKTpr_sd is null and sttDKThiDuapr_sd in (select sttDKThiDuapr from tblDangKyThiDua where duyetDK=1 and ngayDangKy between '" + HttpContext.Current.Session.GetNgayDauKy() + @"' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + @"'
                     and sttDoiTuongTDKTpr_sd in (select sttDoiTuongTDKTpr from tblDMDoiTuongTDKT where maDonVipr_noicongtac= N'" + maDV + "')) and  loaiDoiTuong=N'Cá nhân' and maKhenThuongpr_sd='" + cboDanhHieuWD1.SelectedValue + "'  and maHinhThucTDKTpr_sd='" + cboHinhThucWD1.SelectedValue +
                     "' )as temp)as temp1 where serachName COLLATE SQL_Latin1_General_CP1_CI_AI LIKE N'" + hdfAlphabet1.Value.ToString() + "%'";
                tab = sqlFun.GetData(sqlQuery);
                tab.Columns.Add("ketQua", typeof(string));
                foreach (DataRow dr in tab.Rows)
                {

                    //nếu trong khoảng thời gian có bị kỹ luật sẽ không được xét
                    if (ntsLibraryFunctions.kiemTraDieuKienKhongXetTDKT(dr["sttDoiTuongTDKTpr_sd"].ToString(), dr["maKhenThuongpr_sd"].ToString()) == true)
                    {
                        dr["ketQua"] = "Không đủ điều kiện";
                        tab.AcceptChanges();
                        continue;
                    }
                    string ngayDauKy = HttpContext.Current.Session.GetNgayDauKy().Substring(3, 2) + "/" + HttpContext.Current.Session.GetNgayDauKy().Substring(0, 2) + "/" + HttpContext.Current.Session.GetNgayDauKy().Substring(6, 4);
                    string ngayCuoiKy = HttpContext.Current.Session.GetNgayCuoiKy().Substring(3, 2) + "/" + HttpContext.Current.Session.GetNgayCuoiKy().Substring(0, 2) + "/" + HttpContext.Current.Session.GetNgayCuoiKy().Substring(6, 4);
                    decimal slSKKN = ntsLibraryFunctions.kiemTraSKKN(dr["maKhenThuongpr_sd"].ToString(), dr["sttDoiTuongTDKTpr_sd"].ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                    decimal slDHTD = ntsLibraryFunctions.kiemTraDKBatBuoc(dr["maKhenThuongpr_sd"].ToString(), dr["sttDoiTuongTDKTpr_sd"].ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                    //decimal slDHTDPhu=0;//= ntsLibraryFunctions.kiemTraDKBatBuoc(dr["maKhenThuongpr_sd"].ToString(), dr["sttDoiTuongTDKTpr_sd"].ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);

                    //if (slDHTD == 1) { slDHTDPhu = 1; } else { slDHTDPhu = 0; }
                    decimal slDHTDPhu = ntsLibraryFunctions.kiemTraDKPhu(dr["maKhenThuongpr_sd"].ToString(), dr["sttDoiTuongTDKTpr_sd"].ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                    if ((slSKKN + slDHTD + slDHTDPhu) >= 2)
                    {
                        dr["ketQua"] = "Đủ điều kiện";
                        tab.AcceptChanges();
                    }
                    else
                    {
                        dr["ketQua"] = "Không đủ điều kiện";
                        tab.AcceptChanges();
                    }
                }
                return tab;
            }
        }

        [AjaxPro.AjaxMethod]
        public string KiemTraDoiTuongTDKT(string sttDoiTuongTDKT, string maKhenThuong)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string ketqua = "";
                //nếu trong cấu hình cho phép nhập nhiều lần thì kiểm tra hoặc trong cấu hình chưa có thì kiểm tra
                if (sqlFun.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + maKhenThuong + "' ") == "0" || sqlFun.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + maKhenThuong + "' ") == "")
                {
                    string sql = @"SELECT soQuyetDinh,maKhenThuongpr_sd,capdo=(SELECT capDo FROM dbo.tblDMKhenThuong WHERE maKhenThuongpr = maKhenThuongpr_sd),laThiDua=(SELECT laThiDua FROM dbo.tblDMKhenThuong WHERE maKhenThuongpr = maKhenThuongpr_sd) FROM dbo.tblThiDuaKhenThuong WHERE ngayKyQD BETWEEN '" + HttpContext.Current.Session.GetNgayDauKy() + "' AND '" + HttpContext.Current.Session.GetNgayCuoiKy() + "' AND sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + @"' AND maKhenThuongpr_sd IN (SELECT maKhenThuongpr FROM dbo.tblDMKhenThuong WHERE ngungSD = 0)";
                    DataTable tblThiDuaKhenThuong = new DataTable();
                    tblThiDuaKhenThuong = sqlFun.GetData(sql);
                    string sqlQuery_ = "SELECT convert(nvarchar(50),capDo) as capDo FROM dbo.tblDMKhenThuong WHERE maKhenThuongpr=N'" + maKhenThuong.ToString() + "' AND laThiDua='0'";
                    if (tblThiDuaKhenThuong.Rows.Count > 0)
                    {
                        for (int i = 0; i < tblThiDuaKhenThuong.Rows.Count; i++)
                        {
                            if (tblThiDuaKhenThuong.Rows[i]["maKhenThuongpr_sd"].ToString() == maKhenThuong.ToString())
                            {
                                ketqua = tblThiDuaKhenThuong.Rows[i]["soQuyetDinh"].ToString();
                                break;
                            }
                            if (tblThiDuaKhenThuong.Rows[i]["laThiDua"].ToString() == "False")
                            {
                                if (tblThiDuaKhenThuong.Rows[i]["capDo"].ToString() == sqlFun.GetOneStringField(sqlQuery_))
                                {
                                    ketqua = tblThiDuaKhenThuong.Rows[i]["soQuyetDinh"].ToString();
                                    break;
                                }
                            }
                        }
                        return ketqua;

                    }
                    else
                        return ketqua = "";
                }
                else
                {
                    return ketqua = "";
                }

            }
            catch (Exception)
            {
                return "Loi";
            }
        }
        void AsyncFileUpload_UploadedComplete(object sender, AsyncFileUploadEventArgs e)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string path = "";
                string fileName = "";
                string[] arr = e.filename.Split('\\');
                if (arr.Length > 0)
                    fileName = arr[arr.Length - 1];
                string strDate = DateTime.Now.ToString("ddmmyyyyhhmmss");
                string fileExtension = Path.GetExtension(fileName).Replace(".", "");
                if (Path.GetExtension(fileName).Contains(".jpg") == true || Path.GetExtension(fileName).Contains(".gif") == true || Path.GetExtension(fileName).Contains(".png") == true || Path.GetExtension(fileName).Contains(".bmp") == true)
                {
                    fileName = "hinhAnh_" + HttpContext.Current.Session.GetMaDonVi() + "_" + HttpContext.Current.Session.GetCurrentUserID() + "_" + strDate + "." + fileExtension;
                    if (!System.IO.Directory.Exists(Server.MapPath("~/ThiDuaKhenThuong/HinhAnhTDKT/" + HttpContext.Current.Session.getNamSudung() + "/")))
                    {
                        System.IO.Directory.CreateDirectory(Server.MapPath("~/ThiDuaKhenThuong/HinhAnhTDKT/" + HttpContext.Current.Session.getNamSudung() + "/"));
                    }
                    path = string.Concat(Server.MapPath("~/ThiDuaKhenThuong/HinhAnhTDKT/" + HttpContext.Current.Session.getNamSudung() + "/" + fileName));
                    AsyncFileUpload.SaveAs(path);
                    sqlFun.ExeCuteNonQuery("INSERT INTO dbo.tblHinhAnhKemTheo(sttTDKTpr_sd,duongDan,maDonVipr_sd,ngayThaoTac,nguoiThaoTac) VALUES  (N'" + hdfSTTTDKT.Value.ToString() + "',N'" + "/ThiDuaKhenThuong/HinhAnhTDKT/" + HttpContext.Current.Session.getNamSudung() + "/" + fileName + "',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "')");

                }
            }
            catch
            {
            }
            AsyncFileUpload.ClearState();
            AsyncFileUpload.Dispose();
        }
        protected void AsyncFileUpload_UploadedFileError(object sender, AsyncFileUploadEventArgs e)
        {
        }
        [AjaxPro.AjaxMethod]
        public bool capNhatNgayDK(object[] param)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                SqlConnection sqlCon = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                sqlCon.Open();
                string sql = @"UPDATE tblThiDuaKhenThuong 
                            set  sttDoiTuongTDKTpr_cha=@sttDoiTuongTDKTpr_cha,thamNienQD=@thamNienQD,ghiChu=@ghiChu where sttTDKTpr=@sttTDKTpr";
                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@sttTDKTpr", param[1].ToString()));
                if (string.IsNullOrEmpty(param[2].ToString()))
                {
                    cmd.Parameters.Add(new SqlParameter("@sttDoiTuongTDKTpr_cha", DBNull.Value));
                }
                else
                {
                    cmd.Parameters.Add(new SqlParameter("@sttDoiTuongTDKTpr_cha", param[2].ToString()));
                }
                cmd.Parameters.Add(new SqlParameter("@thamNienQD", param[3].ToString()));
                cmd.Parameters.Add(new SqlParameter("@ghiChu", param[4].ToString()));
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        [AjaxPro.AjaxMethod]
        public bool KTraFile()
        {
            bool kq = true;
            return kq;
        }
        #region "Upload excel"
        void AsyncFileUpload2_UploadedComplete(object sender, AsyncFileUploadEventArgs e)
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

                if (!System.IO.Directory.Exists(Server.MapPath("~/Excel/TDKT/" + HttpContext.Current.Session.GetMaDonVi() + "")))
                {
                    System.IO.Directory.CreateDirectory(Server.MapPath("~/Excel/TDKT/" + HttpContext.Current.Session.GetMaDonVi() + ""));
                }
                path = string.Concat(Server.MapPath("~/Excel/TDKT/" + HttpContext.Current.Session.GetMaDonVi() + "/" + fileName));

                if (Path.GetExtension(fileName).Contains(".xls") == false)
                {
                    return;
                }

                if (!System.IO.File.Exists(path))
                {

                    //System.IO.File.Delete(path);
                    AsyncFileUpload2.SaveAs(path);

                    urlFile = "~/Excel/TDKT/" + HttpContext.Current.Session.GetMaDonVi() + "/" + fileName + "";

                    if (Session["ssGetUrlFile"] == null)
                    {
                        Session.Add("ssGetUrlFile", urlFile);
                    }
                    else
                    {
                        Session["ssGetUrlFile"] = "";
                        Session["ssGetUrlFile"] = urlFile;
                    }
                    //NhapDuLieuFileExcel();
                    //return;
                }
                else
                {
                    //Xoa file neu da ton tai                   
                    System.IO.File.Delete(path);
                    AsyncFileUpload2_UploadedComplete(sender, e);
                }
            }
            catch
            {
            }
            AsyncFileUpload2.ClearState();
            AsyncFileUpload2.Dispose();
        }
        void AsyncFileUpload2_UploadedFileError(object sender, AsyncFileUploadEventArgs e)
        {
        }
        [AjaxPro.AjaxMethod]
        public string getUrlImg()
        {
            if (Session["ssGetUrlFile"] == null)
            {
                return "";
            }
            else
            {
                return Session["ssGetUrlFile"].ToString();
            }
        }
        #endregion NHẬP DỮ LIỆU EXCEL
        [AjaxPro.AjaxMethod]
        public string NhapDuLieuFileExcel()
        {
            string path = "";
            Session["ssloinhapexcelcanhan"] = "";
            decimal sodongthemthanhcong = 0;
            decimal sodongphongban = 0;
            decimal sodongcanhan = 0;
            decimal sodongtong = 0;
            decimal sodongthemkhongthanhcong = 0;
            Session["sodongthemkhongthanhcong"] = "";
            string ssloinhapexcelcanhan = "";
            if (Session["ssGetUrlFile"] != null)
            {
                path = Session["ssGetUrlFile"].ToString();
            }
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
                string sqlSelect = "select * from [KhenThuong$A8:Q20000]";
                OleDbCommand conmand = new OleDbCommand(sqlSelect, connection);
                OleDbDataAdapter oleDA = new OleDbDataAdapter(conmand);
                oleDA.Fill(dtTemp);
                connection.Close();

                //sodongtong = dtTemp.Rows.Count.ToString();
                int sodong = 0, n = 0;
                string sqlInsert = "", sttDoiTuong = "", sttQD = "", hoTen = "", chucVu = "", namSinh = "", donVi = "", noiCongTac = "", soQuyetDinh = "", ngayKhenThuong = "", ngayKyQD = "", danhHieu = "", hinhThuc = "", capQuyetDinh = "", loaiDoiTuong = "", nhapExcel = "", linhVuc = "", diaChi = "", ghiChu = "", maLoaiDoiTuongpr_sd = "";
                string maDV = HttpContext.Current.Session.GetMaDonVi();
                DataTable tblDonViTrucThuoc = new DataTable();
                string strWhereDV = "maDonVipr IN (SELECT maDonvipr FROM dbo.tblDMDonvi WHERE (maDonvipr= N'" + maDV + "' OR (maDonvipr_cha= N'" + maDV + "') OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "')  OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "'))) OR  maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "') ) ) )";
                string sqlDonViTrucThuoc = " select maDonVipr,tenDonVi from tbldmDonVi where " + strWhereDV;
                tblDonViTrucThuoc = _class.GetData(sqlDonViTrucThuoc);
                if (dtTemp.Columns.Count > 0)
                {
                    for (int i = 0; i < dtTemp.Rows.Count; i++)
                    {
                        if (dtTemp.Rows[i][2].ToString() != "")
                        {
                            sodongtong = sodongtong + 1;
                            hoTen = dtTemp.Rows[i][2].ToString();
                            bool loi = false;
                            bool loi_danhxung = false;

                            string[] danhXung = { "Ông", "Bà", "ông", "bà", "ba", "Ong", "Ba", "ong", "Mrs", "Mr", "mrs", "mr" };
                            for (int y = 0; y < danhXung.Length; y++)
                            {
                                if (hoTen.IndexOf(danhXung[y].ToString()) == 0)
                                {
                                    loi_danhxung = true;
                                    break;
                                }
                            }
                            if (loi_danhxung)
                            {
                                sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có họ và tên chứa danh xưng vui lòng kiểm tra lại.</br>";
                                continue;
                            }
                            string[] dauDacBiet = { ".", "@", "#", "$", "%", "^", "&", "*", "=" };
                            for (int y = 0; y < dauDacBiet.Length; y++)
                            {
                                //   int h=tenDoiTuongTDKT.IndexOf(chuoi[y].ToString());
                                if (hoTen.IndexOf(dauDacBiet[y].ToString()) >= 0)
                                {
                                    loi = true;
                                    break;
                                }
                            }
                            if (loi)
                            {
                                sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có họ và tên chứa ký tự không hợp lệ.</br>";
                                continue;
                            }

                            chucVu = dtTemp.Rows[i][3].ToString();
                            namSinh = dtTemp.Rows[i][4].ToString();
                            if (namSinh != "")
                            {
                                if (kiemtrangaythangnam(namSinh) == false)
                                {
                                    sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có ngày sinh (" + dtTemp.Rows[i][4].ToString() + ") không đúng định dạng dd/mm/yyyy hoặc yyyy. </br>";
                                    continue;

                                }
                                else
                                {
                                    if (namSinh.Length == 4)
                                    {
                                        string ngay = "01/01/" + namSinh.ToString().Trim();
                                        namSinh = DateTime.Parse(ngay.Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                    }
                                    else
                                    {
                                        try
                                        {
                                            namSinh = DateTime.Parse(namSinh.ToString().Trim().Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                        }
                                        catch
                                        {
                                            sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có ngày sinh (" + dtTemp.Rows[i][4].ToString() + ") không đúng định dạng dd/mm/yyyy  hoặc yyyy. </br>";
                                            continue;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                namSinh = "";
                            }
                            if (dtTemp.Rows[i][6].ToString() == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " không được để trống đơn vị.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            else
                            {
                                //kiểm tra trong đơn vị có mấy đơn vị trùng tên

                                decimal soluongdonvi = _class.GetOneDecimalField(@"select convert(decimal(18,0),count(maDonvipr)) from tbldmDonVi where tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "' and " + strWhereDV + "");
                                if (soluongdonvi > 1)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Hệ thống có " + soluongdonvi + " đơn vị trực thuộc cùng tên: '" + dtTemp.Rows[i][6].ToString() + "' .Vui lòng kiểm tra lại.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    break;
                                }
                                //lấy mã đơn vị
                                donVi = _class.GetOneStringField("SELECT top 1 maDonvipr FROM dbo.tblDMDonvi WHERE tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "' and " + strWhereDV + "").ToString();
                                if (donVi.ToString() == "")
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có đơn vị '" + dtTemp.Rows[i][6].ToString() + "' không tồn tại trong hệ thống.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {
                                    string timthay = "";
                                    for (int j = 0; j < tblDonViTrucThuoc.Rows.Count; j++)
                                    {
                                        if (donVi.ToString() == tblDonViTrucThuoc.Rows[j]["maDonvipr"].ToString())
                                        {
                                            timthay = "timthay";
                                            break;
                                        }

                                    }
                                    if (timthay.ToString() != "timthay")
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có đơn vị '" + dtTemp.Rows[i][6].ToString() + "' không tồn tại trong danh sách đơn vị trực thuộc.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }
                            soQuyetDinh = dtTemp.Rows[i][7].ToString();
                            if (soQuyetDinh == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " không được để trống số quyết định.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            ngayKhenThuong = dtTemp.Rows[i][9].ToString();
                            if (ngayKhenThuong == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " năm xét khen thưởng không được để trống.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            else
                            {
                                if (kiemtrangaythangnam(ngayKhenThuong) == false || ngayKhenThuong.Length != 4)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " năm xét khen thưởng (" + dtTemp.Rows[i][9].ToString() + ") không đúng định dạng yyyy.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {

                                    try
                                    {

                                        ngayKhenThuong = dtTemp.Rows[i][9].ToString();
                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có năm xét khen thưởng (" + dtTemp.Rows[i][9].ToString() + ") không đúng định dạng yyyy.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }
                            ngayKyQD = dtTemp.Rows[i][8].ToString();
                            if (ngayKyQD == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " ngày ký quyết định không được để trống.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            else
                            {
                                if (kiemtrangaythangnam(ngayKyQD) == false)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " ngày ký quyết định (" + dtTemp.Rows[i][8].ToString() + ") không đúng định dạng dd/mm/yyyy.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {

                                    try
                                    {

                                        ngayKyQD = DateTime.Parse(dtTemp.Rows[i][8].ToString().Trim().Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có ngày ký quyết định (" + dtTemp.Rows[i][8].ToString() + ") không đúng định dạng dd/mm/yyyy.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }

                            DataTable tblDanhHieu_HinhThuc = new DataTable();
                            tblDanhHieu_HinhThuc = _class.GetData("SELECT top 1 maKhenThuongpr,capDo,laThiDua FROM dbo.tblDMKhenThuong WHERE loaiDoiTuong LIKE N'%Cá nhân%' and tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");

                            if (tblDanhHieu_HinhThuc.Rows.Count > 0)//làm tới đây
                            {
                                danhHieu = tblDanhHieu_HinhThuc.Rows[0]["maKhenThuongpr"].ToString();
                            }
                            else
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " danh hiệu '" + dtTemp.Rows[i][10] + "' không tồn tại trong dữ liệu.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }

                            hinhThuc = _class.GetOneStringField("SELECT maHinhThucTDKTpr FROM dbo.tblDMHinhThucTDKT WHERE tenHinhThucTDKT =N'" + dtTemp.Rows[i][11] + "'");
                            if (hinhThuc == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có hình thức khen thưởng '" + dtTemp.Rows[i][11] + "' không tồn tại trong dữ liệu.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }

                            linhVuc = _class.GetOneStringField("SELECT maLinhVucpr FROM tblDMLinhVuc WHERE tenLinhVuc=N'" + dtTemp.Rows[i][13].ToString() + "'");
                            if (linhVuc == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " lĩnh vực không tồn tại trong hệ thống.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }

                            maLoaiDoiTuongpr_sd = _class.GetOneStringField("SELECT maLoaiDoiTuongpr FROM tblDMLoaiDoiTuong WHERE tenLoaiDoiTuong=N'" + dtTemp.Rows[i][16].ToString() + "'");
                            if (maLoaiDoiTuongpr_sd == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có loại đối tượng không tồn tại trong hệ thống.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }

                            capQuyetDinh = _class.GetOneStringField("SELECT top 1 convert(nvarchar(50),capDo) FROM dbo.tblDMKhenThuong WHERE tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");
                            diaChi = dtTemp.Rows[i][14].ToString();
                            ghiChu = dtTemp.Rows[i][15].ToString();
                            loaiDoiTuong = "Cá nhân";
                            nhapExcel = "1";
                            string sttdoituongtapthe = "";
                            DataTable tableDoiTuong = new DataTable();
                            //kiểm tra đối tượng
                            string check_doituong = "SELECT top 1 sttDoiTuongTDKTpr,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE tenDoiTuongTDKT = N'" + hoTen + "' and loaiDoiTuong =N'Cá nhân' " + (namSinh == "" ? "" : "and ngaySinh ='" + namSinh + "' ") + " and maDonVipr_noicongtac = '" + donVi + "'";
                            tableDoiTuong = _class.GetData(check_doituong);
                            if (tableDoiTuong.Rows.Count > 0)
                            {
                                sttDoiTuong = tableDoiTuong.Rows[0]["sttDoiTuongTDKTpr"].ToString();

                                //kiểm tra đối tượng tập thể
                                DataTable tableDoiTuong_TapThe = new DataTable();
                                string check_doituong_tapthe = "SELECT top 1 sttDoiTuongTDKTpr,tenChucVuQL,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE  tenDoiTuongTDKT = N'" + dtTemp.Rows[i][5] + "' AND maDonVipr_noicongtac = '" + donVi + "' and loaiDoiTuong =N'Tập thể'";
                                tableDoiTuong_TapThe = _class.GetData(check_doituong_tapthe);
                                if (tableDoiTuong_TapThe.Rows.Count > 0)
                                {
                                    sttdoituongtapthe = tableDoiTuong_TapThe.Rows[0]["sttDoiTuongTDKTpr"].ToString();
                                }
                                else
                                {
                                    if (dtTemp.Rows[i][5].ToString() != "")
                                    {
                                        SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                                        SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                        string soPhieu = "";
                                        string sql = "SELECT MAX(CONVERT(DECIMAL,RIGHT(maDoiTuong,6))) FROM dbo.tblDMDoiTuongTDKT WHERE maDoiTuong like N'" + HttpContext.Current.Session.GetMaDonVi() + "-%' and ISNUMERIC(RIGHT(maDoiTuong,6)) = 1 AND maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                        decimal _vNewKey = sqlFun.GetOneDecimalField(sql) + 1;
                                        soPhieu = HttpContext.Current.Session.GetMaDonVi() + "-" + _vNewKey.ToString("000000");
                                        sqlConn.Open();
                                        SqlCommand sqlCom = new SqlCommand("INSERT dbo.tblDMDoiTuongTDKT (maDoiTuong,tenDoiTuongTDKT,loaiDoiTuong,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,maDonVipr_noicongtac,nhapExcel)"
                                       + "  VALUES(@soPhieu,@hoTen,@loaiDoiTuong,@maDonVipr_sd,@ngayThaoTac,@nguoiThaoTac,@maDonVipr_noicongtac,@nhapExcel)", sqlConn);
                                        sqlCom.Parameters.Add(new SqlParameter("@soPhieu", soPhieu));
                                        sqlCom.Parameters.Add(new SqlParameter("@hoTen", dtTemp.Rows[i][5].ToString()));
                                        sqlCom.Parameters.Add(new SqlParameter("@loaiDoiTuong", "Tập thể"));
                                        sqlCom.Parameters.Add(new SqlParameter("@maDonVipr_sd", HttpContext.Current.Session.GetMaDonVi()));
                                        sqlCom.Parameters.Add(new SqlParameter("@ngayThaoTac", HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy")));
                                        sqlCom.Parameters.Add(new SqlParameter("@nguoiThaoTac", HttpContext.Current.Session.GetCurrentUserID()));
                                        sqlCom.Parameters.Add(new SqlParameter("@maDonVipr_noicongtac", donVi));
                                        sqlCom.Parameters.Add(new SqlParameter("@nhapExcel", nhapExcel));
                                        if (sqlCom.ExecuteNonQuery() > 0)
                                        {
                                            decimal sttDoiTuong_ = sqlFun.GetOneDecimalField("SELECT MAX(CONVERT(DECIMAL,sttDoiTuongTDKTpr)) FROM dbo.tblDMDoiTuongTDKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'");
                                            sttdoituongtapthe = sttDoiTuong_.ToString();
                                            sodongphongban = sodongphongban + 1;
                                            sqlConn.Close();
                                        }
                                        else
                                        {
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " thêm mới phòng ban '" + dtTemp.Rows[i][5].ToString() + "' không thành công.</br>";
                                            sodongthemkhongthanhcong += 1;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        sttdoituongtapthe = "";
                                    }

                                }
                            }
                            else
                            {
                                DataTable tableDoiTuong_TapThe = new DataTable();
                                //kiểm tra đối tượng tập thể

                                string check_doituong_tapthe = "SELECT top 1 sttDoiTuongTDKTpr,tenChucVuQL,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE  tenDoiTuongTDKT = N'" + dtTemp.Rows[i][5] + "' AND maDonVipr_noicongtac = '" + donVi + "' and loaiDoiTuong =N'Tập thể'";
                                tableDoiTuong_TapThe = _class.GetData(check_doituong_tapthe);
                                if (tableDoiTuong_TapThe.Rows.Count > 0)
                                {
                                    sttdoituongtapthe = tableDoiTuong_TapThe.Rows[0]["sttDoiTuongTDKTpr"].ToString();
                                }
                                else
                                {
                                    if (dtTemp.Rows[i][5].ToString() != "")
                                    {
                                        SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                                        SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                        string soPhieu = "";
                                        string sql = "SELECT MAX(CONVERT(DECIMAL,RIGHT(maDoiTuong,6))) FROM dbo.tblDMDoiTuongTDKT WHERE maDoiTuong like N'" + HttpContext.Current.Session.GetMaDonVi() + "-%' and ISNUMERIC(RIGHT(maDoiTuong,6)) = 1 AND maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                        decimal _vNewKey = sqlFun.GetOneDecimalField(sql) + 1;
                                        soPhieu = HttpContext.Current.Session.GetMaDonVi() + "-" + _vNewKey.ToString("000000");
                                        sqlConn.Open();
                                        SqlCommand sqlCom = new SqlCommand("INSERT dbo.tblDMDoiTuongTDKT (maDoiTuong,tenDoiTuongTDKT,loaiDoiTuong,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,maDonVipr_noicongtac,nhapExcel)"
                                       + "  VALUES(@soPhieu,@hoTen,@loaiDoiTuong,@maDonVipr_sd,@ngayThaoTac,@nguoiThaoTac,@maDonVipr_noicongtac,@nhapExcel)", sqlConn);
                                        sqlCom.Parameters.Add(new SqlParameter("@soPhieu", soPhieu));
                                        sqlCom.Parameters.Add(new SqlParameter("@hoTen", dtTemp.Rows[i][5].ToString()));
                                        sqlCom.Parameters.Add(new SqlParameter("@loaiDoiTuong", "Tập thể"));
                                        sqlCom.Parameters.Add(new SqlParameter("@maDonVipr_sd", HttpContext.Current.Session.GetMaDonVi()));
                                        sqlCom.Parameters.Add(new SqlParameter("@ngayThaoTac", HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy")));
                                        sqlCom.Parameters.Add(new SqlParameter("@nguoiThaoTac", HttpContext.Current.Session.GetCurrentUserID()));
                                        sqlCom.Parameters.Add(new SqlParameter("@maDonVipr_noicongtac", donVi));
                                        sqlCom.Parameters.Add(new SqlParameter("@nhapExcel", nhapExcel));
                                        if (sqlCom.ExecuteNonQuery() > 0)
                                        {
                                            decimal sttDoiTuong_ = sqlFun.GetOneDecimalField("SELECT MAX(CONVERT(DECIMAL,sttDoiTuongTDKTpr)) FROM dbo.tblDMDoiTuongTDKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'");
                                            sttdoituongtapthe = sttDoiTuong_.ToString();
                                            sodongphongban = sodongphongban + 1;
                                            sqlConn.Close();
                                        }
                                        else
                                        {
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " thêm mới phòng ban '" + dtTemp.Rows[i][5].ToString() + "' không thành công.</br>";
                                            sodongthemkhongthanhcong += 1;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        sttdoituongtapthe = "";
                                    }

                                }
                                SqlConnection sqlConn_ = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                                SqlFunction sqlFun_ = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                string soPhieu_ = "";
                                string sql_ = "SELECT MAX(CONVERT(DECIMAL,RIGHT(maDoiTuong,6))) FROM dbo.tblDMDoiTuongTDKT WHERE maDoiTuong like N'" + HttpContext.Current.Session.GetMaDonVi() + "-%' and ISNUMERIC(RIGHT(maDoiTuong,6)) = 1 AND maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                decimal _vNewKey_ = sqlFun_.GetOneDecimalField(sql_) + 1;
                                soPhieu_ = HttpContext.Current.Session.GetMaDonVi() + "-" + _vNewKey_.ToString("000000");
                                sqlConn_.Open();
                                SqlCommand sqlCom_ = new SqlCommand("INSERT dbo.tblDMDoiTuongTDKT (sttDoiTuongTDKTpr_cha,danhXungpr_sd,namSinh,maDoiTuong,tenDoiTuongTDKT,tenChucVuQL,ngaySinh,loaiDoiTuong,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,maDonVipr_noicongtac,nhapExcel,diaChi,maLoaiDoiTuongpr_sd)"
                               + "  VALUES(@sttDoiTuongTDKTpr_cha,@danhXungpr_sd,@namSinh,@soPhieu,@hoTen,@chucVu,@ngaySinh,@loaiDoiTuong,@maDonVipr_sd,@ngayThaoTac,@nguoiThaoTac,@maDonVipr_noicongtac,@nhapExcel,@diaChi,@maLoaiDoiTuongpr_sd)", sqlConn_);
                                sqlCom_.Parameters.Add(new SqlParameter("@soPhieu", soPhieu_));
                                if (string.IsNullOrEmpty(sttdoituongtapthe.ToString()))
                                {
                                    sqlCom_.Parameters.Add(new SqlParameter("@sttDoiTuongTDKTpr_cha", DBNull.Value));
                                }
                                else
                                {
                                    sqlCom_.Parameters.Add(new SqlParameter("@sttDoiTuongTDKTpr_cha", sttdoituongtapthe));
                                }
                                sqlCom_.Parameters.Add(new SqlParameter("@hoTen", hoTen));
                                sqlCom_.Parameters.Add(new SqlParameter("@chucVu", chucVu));
                                if (string.IsNullOrEmpty(namSinh.ToString()))
                                {
                                    sqlCom_.Parameters.Add(new SqlParameter("@ngaySinh", DBNull.Value));
                                }
                                else
                                {
                                    sqlCom_.Parameters.Add(new SqlParameter("@ngaySinh", namSinh));
                                }
                                if (dtTemp.Rows[i][4].ToString().Length == 4)
                                {
                                    sqlCom_.Parameters.Add(new SqlParameter("@namSinh", "1"));
                                }
                                else
                                {
                                    sqlCom_.Parameters.Add(new SqlParameter("@namSinh", "0"));
                                }
                                sqlCom_.Parameters.Add(new SqlParameter("@loaiDoiTuong", loaiDoiTuong));
                                sqlCom_.Parameters.Add(new SqlParameter("@maDonVipr_sd", HttpContext.Current.Session.GetMaDonVi()));
                                sqlCom_.Parameters.Add(new SqlParameter("@ngayThaoTac", HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy")));
                                sqlCom_.Parameters.Add(new SqlParameter("@nguoiThaoTac", HttpContext.Current.Session.GetCurrentUserID()));
                                sqlCom_.Parameters.Add(new SqlParameter("@maDonVipr_noicongtac", donVi));
                                sqlCom_.Parameters.Add(new SqlParameter("@nhapExcel", nhapExcel));
                                sqlCom_.Parameters.Add(new SqlParameter("@danhXungpr_sd", dtTemp.Rows[i][1].ToString().Trim()));
                                sqlCom_.Parameters.Add(new SqlParameter("@diaChi", diaChi));
                                sqlCom_.Parameters.Add(new SqlParameter("@maLoaiDoiTuongpr_sd", maLoaiDoiTuongpr_sd));
                                if (sqlCom_.ExecuteNonQuery() > 0)
                                {
                                    decimal sttDoiTuong_ = sqlFun_.GetOneDecimalField("SELECT MAX(sttDoiTuongTDKTpr) FROM dbo.tblDMDoiTuongTDKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'");
                                    sttDoiTuong = sttDoiTuong_.ToString();
                                    sodongcanhan = sodongcanhan + 1;
                                    sqlConn_.Close();
                                }
                                else
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " thêm mới không thành công.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                            }
                            //nếu trong cấu hình cho phép nhập nhiều lần thì kiểm tra hoặc trong cấu hình chưa có thì kiểm tra
                            if (_class.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + danhHieu + "' ") == "0" || _class.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + danhHieu + "' ") == "")
                            {
                                //kiểm tra hình thức khen thưởng của một cấp đã được nhập mấy lần rồi
                                if (_class.CheckHasRecord(@"select sttDoiTuongTDKTpr_sd from tblThiDuaKhenThuong where sttDoiTuongTDKTpr_sd ='" + sttDoiTuong + "' and maKhenThuongpr_sd in (select maKhenThuongpr from tblDMKhenThuong where capDo = N'" + tblDanhHieu_HinhThuc.Rows[0]["capDo"].ToString() + "' and laThiDua ='0' and ngungSD ='0') and  year(ngayKyQD) = '" + ngayKyQD.Split('/')[2].ToString() + "' ") == true)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " hình thức khen thưởng '" + dtTemp.Rows[i][10] + "' chỉ được phép nhập một lần cho một cấp trong năm.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                            }
                            //nếu trong cấu hình cho phép nhập nhiều lần thì kiểm tra hoặc trong cấu hình chưa có thì kiểm tra
                            if (_class.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + danhHieu + "' ") == "0" || _class.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + danhHieu + "' ") == "")
                            {
                                //kiểm tra đối tượng đã được nhập khen thưởng hay chưa
                                DataTable tblThiDuaKhenThuong = new DataTable();
                                tblThiDuaKhenThuong = _class.GetData(@"select top 1 sttDoiTuongTDKTpr_sd,soQuyetDinh,CONVERT(NVARCHAR(50),ngayKyQD,103) AS ngayKyQD from tblThiDuaKhenThuong where sttDoiTuongTDKTpr_sd ='" + sttDoiTuong + "' and year(ngayKyQD) = '" + ngayKyQD.Split('/')[2].ToString() + "' and maKhenThuongpr_sd = '" + danhHieu + "'");
                                if (tblThiDuaKhenThuong.Rows.Count > 0)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " đã được nhập khen thưởng số '" + tblThiDuaKhenThuong.Rows[0]["soQuyetDinh"].ToString() + "' ngày quyết định '" + tblThiDuaKhenThuong.Rows[0]["ngayKyQD"].ToString() + "' rồi.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {
                                    //kiểm tra quyết định
                                    DataTable tableQuyetDinh = new DataTable();

                                    string stringQuyetDinh = "SELECT sttQuyetDinhKTpr FROM dbo.tblQuyetDinhKT WHERE maDonVipr_sd =N'" + HttpContext.Current.Session.GetMaDonVi() + "' AND maKhenThuongpr_sd =N'" + danhHieu + "' AND maHinhThucTDKTpr_sd =N'" + hinhThuc + "' AND soQuyetDinh=N'" + soQuyetDinh + "' AND ngayKy=N'" + ngayKyQD + "' AND capQuyetDinh=N'" + capQuyetDinh + "' AND loaiDoiTuong=N'" + loaiDoiTuong + "'";
                                    tableQuyetDinh = _class.GetData(stringQuyetDinh);
                                    if (tableQuyetDinh.Rows.Count > 0)
                                    {
                                        sttQD = tableQuyetDinh.Rows[0]["sttQuyetDinhKTpr"].ToString();
                                    }
                                    else
                                    {
                                        try
                                        {
                                            string sqlInsert_ = @"INSERT INTO [dbo].[tblQuyetDinhKT] ([noiDung],[ngayKhenThuong], [maKhenThuongpr_sd], [maHinhThucTDKTpr_sd], [soQuyetDinh], [ngayKy], [capQuyetDinh], [ngayThaoTac], [nguoiThaoTac], [maDonVipr_sd],[loaiDoiTuong],[nhapExcel],[namXetKhenThuong],maLinhVucpr_sd) values
	                            ( N'" + dtTemp.Rows[i][12].ToString() + "', N'" + ngayKyQD + "', N'" + danhHieu + "', N'" + hinhThuc + "', N'" + soQuyetDinh + "',N'" + ngayKyQD + "',N'" + capQuyetDinh + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + loaiDoiTuong + "','1', N'" + ngayKhenThuong.ToString() + "',N'" + linhVuc + "')";
                                            if (_class.ExeCuteNonQuery(sqlInsert_) == true)
                                            {
                                                string sqlchuoi = "SELECT CONVERT(DECIMAL,MAX(sttQuyetDinhKTpr)) FROM dbo.tblQuyetDinhKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                                decimal sttQD_ = _class.GetOneDecimalField(sqlchuoi);
                                                sttQD = sttQD_.ToString();
                                            }
                                            else
                                            {
                                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                                sodongthemkhongthanhcong += 1;
                                                continue;
                                            }
                                        }
                                        catch
                                        {
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                            sodongthemkhongthanhcong += 1;
                                            continue;
                                        }
                                    }
                                    try
                                    {
                                        //insert đối tượng vào quyết định :tblthiduakhenthuong
                                        if (string.IsNullOrEmpty(sttdoituongtapthe.ToString()))
                                        {
                                            string sqlInsertkt = @"INSERT INTO [dbo].[tblThiDuaKhenThuong] (noiDungTDKT, tinhTrang,soQuyetDinh,ngayKhenThuong,ngayKyQD,capQuyetDinh,loaiDoiTuong,maKhenThuongpr_sd,maHinhThucTDKTpr_sd,sttDoiTuongTDKTpr_sd,sttQuyetDinhKTpr_sd,nhapExcel,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,ngayDangKy,namXetKhenThuong,tenChucVu,maDonVipr_noicongtac,maLinhVucpr_sd,ghiChu) values
	                            ( N'" + dtTemp.Rows[i][12].ToString() + "',  N'Có quyết định', N'" + soQuyetDinh + "', N'" + ngayKyQD + "', N'" + ngayKyQD + "', N'" + capQuyetDinh + "',N'" + loaiDoiTuong + "',N'" + danhHieu + "',N'" + hinhThuc + "',N'" + sttDoiTuong + "',N'" + sttQD + "','1',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',null, N'" + ngayKhenThuong.ToString() + "',N'" + chucVu + "',N'" + donVi + "',N'" + linhVuc + "',N'" + ghiChu + "')";
                                            _class.ExeCuteNonQuery(sqlInsertkt);
                                            sodongthemthanhcong = sodongthemthanhcong + 1;
                                        }
                                        else
                                        {
                                            string sqlInsertkt = @"INSERT INTO [dbo].[tblThiDuaKhenThuong] (noiDungTDKT,sttDoiTuongTDKTpr_cha,tinhTrang,soQuyetDinh,ngayKhenThuong,ngayKyQD,capQuyetDinh,loaiDoiTuong,maKhenThuongpr_sd,maHinhThucTDKTpr_sd,sttDoiTuongTDKTpr_sd,sttQuyetDinhKTpr_sd,nhapExcel,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,ngayDangKy,namXetKhenThuong,tenChucVu,maDonVipr_noicongtac,maLinhVucpr_sd,ghiChu) values
	                            ( N'" + dtTemp.Rows[i][12].ToString() + "', N'" + sttdoituongtapthe + "',N'Có quyết định', N'" + soQuyetDinh + "', N'" + ngayKyQD + "', N'" + ngayKyQD + "', N'" + capQuyetDinh + "',N'" + loaiDoiTuong + "',N'" + danhHieu + "',N'" + hinhThuc + "',N'" + sttDoiTuong + "',N'" + sttQD + "','1',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',null, N'" + ngayKhenThuong.ToString() + "',N'" + chucVu + "',N'" + donVi + "',N'" + linhVuc + "',N'" + ghiChu + "')";
                                            _class.ExeCuteNonQuery(sqlInsertkt);
                                            sodongthemthanhcong = sodongthemthanhcong + 1;
                                        }

                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' khen thưởng cho cá nhân: " + hoTen + " không thành công.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }
                            else
                            {
                                //kiểm tra quyết định
                                DataTable tableQuyetDinh = new DataTable();
                                DataTable tblThiDuaKhenThuong = new DataTable();
                                tblThiDuaKhenThuong = _class.GetData(@"select top 1 sttDoiTuongTDKTpr_sd,soQuyetDinh,CONVERT(NVARCHAR(50),ngayKyQD,103) AS ngayKyQD from tblThiDuaKhenThuong where sttDoiTuongTDKTpr_sd ='" + sttDoiTuong + "' and year(ngayKyQD) = '" + ngayKyQD.Split('/')[2].ToString() + "' and maKhenThuongpr_sd = '" + danhHieu + "'");
                                if (tblThiDuaKhenThuong.Rows.Count > 0)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " đã được nhập khen thưởng số '" + tblThiDuaKhenThuong.Rows[0]["soQuyetDinh"].ToString() + "' ngày quyết định '" + tblThiDuaKhenThuong.Rows[0]["ngayKyQD"].ToString() + "' rồi.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {
                                    string stringQuyetDinh = "SELECT sttQuyetDinhKTpr FROM dbo.tblQuyetDinhKT WHERE maDonVipr_sd =N'" + HttpContext.Current.Session.GetMaDonVi() + "'  AND maKhenThuongpr_sd =N'" + danhHieu + "' AND maHinhThucTDKTpr_sd =N'" + hinhThuc + "' AND soQuyetDinh=N'" + soQuyetDinh + "' AND ngayKy=N'" + ngayKyQD + "' AND capQuyetDinh=N'" + capQuyetDinh + "' AND loaiDoiTuong=N'" + loaiDoiTuong + "'";
                                    tableQuyetDinh = _class.GetData(stringQuyetDinh);
                                    if (tableQuyetDinh.Rows.Count > 0)
                                    {
                                        sttQD = tableQuyetDinh.Rows[0]["sttQuyetDinhKTpr"].ToString();
                                    }
                                    else
                                    {
                                        try
                                        {
                                            string sqlInsert_ = @"INSERT INTO [dbo].[tblQuyetDinhKT] ([noiDung],[ngayKhenThuong], [maKhenThuongpr_sd], [maHinhThucTDKTpr_sd], [soQuyetDinh], [ngayKy], [capQuyetDinh], [ngayThaoTac], [nguoiThaoTac], [maDonVipr_sd],[loaiDoiTuong],[nhapExcel],namXetKhenThuong,maLinhVucpr_sd) values
	                            ( N'" + dtTemp.Rows[i][12].ToString() + "',N'" + ngayKyQD + "', N'" + danhHieu + "', N'" + hinhThuc + "', N'" + soQuyetDinh + "',N'" + ngayKyQD + "',N'" + capQuyetDinh + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + loaiDoiTuong + "','1', N'" + ngayKhenThuong.ToString() + "',N'" + linhVuc + "')";
                                            if (_class.ExeCuteNonQuery(sqlInsert_) == true)
                                            {
                                                string sqlchuoi = "SELECT CONVERT(DECIMAL,MAX(sttQuyetDinhKTpr)) FROM dbo.tblQuyetDinhKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                                decimal sttQD_ = _class.GetOneDecimalField(sqlchuoi);


                                                sttQD = sttQD_.ToString();
                                            }
                                            else
                                            {
                                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                                sodongthemkhongthanhcong += 1;
                                                continue;
                                            }
                                        }
                                        catch
                                        {
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                            sodongthemkhongthanhcong += 1;
                                            continue;
                                        }
                                    }
                                    try
                                    {
                                        //insert đối tượng vào quyết định :tblthiduakhenthuong
                                        if (string.IsNullOrEmpty(sttdoituongtapthe.ToString()))
                                        {
                                            string sqlInsertkt = @"INSERT INTO [dbo].[tblThiDuaKhenThuong] (noiDungTDKT, tinhTrang,soQuyetDinh,ngayKhenThuong,ngayKyQD,capQuyetDinh,loaiDoiTuong,maKhenThuongpr_sd,maHinhThucTDKTpr_sd,sttDoiTuongTDKTpr_sd,sttQuyetDinhKTpr_sd,nhapExcel,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,ngayDangKy,namXetKhenThuong,tenChucVu,maDonVipr_noicongtac,maLinhVucpr_sd,ghiChu) values
	                            (N'" + dtTemp.Rows[i][12].ToString() + "', N'Có quyết định', N'" + soQuyetDinh + "', N'" + ngayKyQD + "', N'" + ngayKyQD + "', N'" + capQuyetDinh + "',N'" + loaiDoiTuong + "',N'" + danhHieu + "',N'" + hinhThuc + "',N'" + sttDoiTuong + "',N'" + sttQD + "','1',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',null, N'" + ngayKhenThuong.ToString() + "',N'" + chucVu + "',N'" + donVi + "',N'" + linhVuc + "',N'" + ghiChu + "')";
                                            _class.ExeCuteNonQuery(sqlInsertkt);
                                            sodongthemthanhcong = sodongthemthanhcong + 1;
                                        }
                                        else
                                        {
                                            string sqlInsertkt = @"INSERT INTO [dbo].[tblThiDuaKhenThuong] (noiDungTDKT,sttDoiTuongTDKTpr_cha,tinhTrang,soQuyetDinh,ngayKhenThuong,ngayKyQD,capQuyetDinh,loaiDoiTuong,maKhenThuongpr_sd,maHinhThucTDKTpr_sd,sttDoiTuongTDKTpr_sd,sttQuyetDinhKTpr_sd,nhapExcel,maDonVipr_sd,ngayThaoTac,nguoiThaoTac,ngayDangKy,namXetKhenThuong,tenChucVu,maDonVipr_noicongtac,maLinhVucpr_sd,ghiChu) values
	                            (N'" + dtTemp.Rows[i][12].ToString() + "',N'" + sttdoituongtapthe + "',N'Có quyết định', N'" + soQuyetDinh + "', N'" + ngayKyQD + "', N'" + ngayKyQD + "', N'" + capQuyetDinh + "',N'" + loaiDoiTuong + "',N'" + danhHieu + "',N'" + hinhThuc + "',N'" + sttDoiTuong + "',N'" + sttQD + "','1',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',null, N'" + ngayKhenThuong.ToString() + "',N'" + chucVu + "',N'" + donVi + "',N'" + linhVuc + "',N'" + ghiChu + "')";
                                            _class.ExeCuteNonQuery(sqlInsertkt);
                                            sodongthemthanhcong = sodongthemthanhcong + 1;
                                        }

                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' khen thưởng cho cá nhân: " + hoTen + " không thành công.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }

                        }
                        else
                        { break; }
                    }
                }

                Session["ssloinhapexcelcanhan"] = ssloinhapexcelcanhan;

                if (sodongthemkhongthanhcong == 0)
                {
                    return "Đã nhận: " + sodongthemthanhcong.ToString() + "/" + sodongtong.ToString() + " dòng. Trong đó thêm mới phòng ban là: " + sodongphongban.ToString() + " phòng ban - Cá nhân: " + sodongcanhan.ToString() + " cá nhân";
                }
                else
                {
                    return "Đã nhận: " + sodongthemthanhcong.ToString() + "/" + sodongtong.ToString() + " dòng.  Trong đó thêm mới phòng ban là: " + sodongphongban.ToString() + " phòng ban - Cá nhân: '" + sodongcanhan.ToString() + "' cá nhân. Có: " + sodongthemkhongthanhcong.ToString() + " dòng bị sai thông tin:<a href='#' onclick='hienThongBaoLoi();return false;'> Bấm xem chi tiết</a>";
                }
            }
            catch (Exception ex)
            {

                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "Tải tập tin excel không thành công. lý do " + ex.Message.ToString() + "</br>";
                return ssloinhapexcelcanhan;
            }
        }
        [AjaxPro.AjaxMethod]
        public string NhapDuLieuFileExcel_DangKy()
        {
            string path = "";
            Session["ssloinhapexcelcanhan"] = "";
            decimal sodongthemthanhcong = 0;
            decimal sodongphongban = 0;
            decimal sodongcanhan = 0;
            decimal sodongtong = 0;
            decimal sodongthemkhongthanhcong = 0;
            Session["sodongthemkhongthanhcong"] = "";
            string ssloinhapexcelcanhan = "";
            if (Session["ssGetUrlFile"] != null)
            {
                path = Session["ssGetUrlFile"].ToString();
            }
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
                string sqlSelect = "select * from [KhenThuong$A8:M2000]";
                OleDbCommand conmand = new OleDbCommand(sqlSelect, connection);
                OleDbDataAdapter oleDA = new OleDbDataAdapter(conmand);
                oleDA.Fill(dtTemp);
                connection.Close();

                //sodongtong = dtTemp.Rows.Count.ToString();
                int sodong = 0, n = 0;
                string sqlInsert = "", sttDoiTuong = "", sttQD = "", hoTen = "", chucVu = "", namSinh = "", donVi = "", noiCongTac = "", soQuyetDinh = "", ngayKhenThuong = "", ngayKyQD = "", danhHieu = "", hinhThuc = "", capQuyetDinh = "", loaiDoiTuong = "", nhapExcel = "", maDonViCongTac = "";
                string maDV = HttpContext.Current.Session.GetMaDonVi();
                DataTable tblDonViTrucThuoc = new DataTable();
                string strWhereDV = "maDonVipr IN (SELECT maDonvipr FROM dbo.tblDMDonvi WHERE (maDonvipr= N'" + maDV + "' OR (maDonvipr_cha= N'" + maDV + "') OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "')  OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "'))) OR  maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "') ) ) )";
                string sqlDonViTrucThuoc = " select maDonVipr,tenDonVi from tbldmDonVi where " + strWhereDV;
                tblDonViTrucThuoc = _class.GetData(sqlDonViTrucThuoc);
                if (dtTemp.Columns.Count > 0)
                {
                    for (int i = 0; i < dtTemp.Rows.Count; i++)
                    {
                        if (dtTemp.Rows[i][2].ToString() != "")
                        {
                            sodongtong = sodongtong + 1;
                            hoTen = dtTemp.Rows[i][2].ToString();
                            bool loi = false;
                            bool loi_danhxung = false;
                            string[] danhXung = { "Ông", "Bà", "ông", "bà", "ba", "Ong", "Ba", "ong", "Mrs", "Mr", "mrs", "mr" };
                            for (int y = 0; y < danhXung.Length; y++)
                            {
                                if (hoTen.IndexOf(danhXung[y].ToString()) == 0)
                                {
                                    loi_danhxung = true;
                                    break;
                                }
                            }
                            if (loi_danhxung)
                            {
                                sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có họ và tên chứa danh xưng vui lòng kiểm tra lại.</br>";
                                continue;
                            }
                            string[] dauDacBiet = { ".", "@", "#", "$", "%", "^", "&", "*", "=" };
                            for (int y = 0; y < dauDacBiet.Length; y++)
                            {

                                if (hoTen.IndexOf(dauDacBiet[y].ToString()) >= 0)
                                {
                                    loi = true;
                                    break;
                                }
                            }
                            if (loi)
                            {
                                sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có họ và tên chứa ký tự không hợp lệ.</br>";
                                continue;
                            }

                            chucVu = dtTemp.Rows[i][3].ToString();
                            namSinh = dtTemp.Rows[i][4].ToString();
                            if (namSinh != "")
                            {
                                if (kiemtrangaythangnam(namSinh) == false)
                                {
                                    sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có ngày sinh (" + dtTemp.Rows[i][4].ToString() + ") không đúng định dạng dd/mm/yyyy hoặc yyyy. </br>";
                                    continue;

                                }
                                else
                                {
                                    if (namSinh.Length == 4)
                                    {
                                        string ngay = "01/01/" + namSinh.ToString().Trim();
                                        namSinh = DateTime.Parse(ngay.Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                    }
                                    else
                                    {
                                        try
                                        {
                                            namSinh = DateTime.Parse(namSinh.ToString().Trim().Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                        }
                                        catch
                                        {
                                            sodongthemkhongthanhcong = sodongthemkhongthanhcong + 1;
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có ngày sinh (" + dtTemp.Rows[i][4].ToString() + ") không đúng định dạng dd/mm/yyyy  hoặc yyyy. </br>";
                                            continue;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                namSinh = "";
                            }
                            if (dtTemp.Rows[i][6].ToString() == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " không được để trống đơn vị.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            else
                            {
                                decimal soluongdonvi = _class.GetOneDecimalField(@"select convert(decimal(18,0),count(maDonvipr)) from tbldmDonVi where tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "' and " + strWhereDV + "");
                                if (soluongdonvi > 1)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Hệ thống có " + soluongdonvi + " đơn vị trực thuộc cùng tên: '" + dtTemp.Rows[i][6].ToString() + "' .Vui lòng kiểm tra lại.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    break;
                                }
                                donVi = _class.GetOneStringField("SELECT top 1 maDonvipr FROM dbo.tblDMDonvi WHERE tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "' and " + strWhereDV + "").ToString();
                                maDonViCongTac = _class.GetOneStringField("SELECT top 1 maDonvipr FROM dbo.tblDMDonvi WHERE tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "'").ToString();
                                if (donVi.ToString() == "")
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có đơn vị '" + dtTemp.Rows[i][6].ToString() + "' không tồn tại trong hệ thống.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {
                                    string timthay = "";
                                    for (int j = 0; j < tblDonViTrucThuoc.Rows.Count; j++)
                                    {
                                        if (donVi.ToString() == tblDonViTrucThuoc.Rows[j]["maDonvipr"].ToString())
                                        {
                                            timthay = "timthay";
                                            break;
                                        }

                                    }
                                    if (timthay.ToString() != "timthay")
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có đơn vị '" + dtTemp.Rows[i][6].ToString() + "' không tồn tại trong danh sách đơn vị trực thuộc.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }
                            soQuyetDinh = dtTemp.Rows[i][7].ToString();
                            if (soQuyetDinh == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " không được để trống số quyết định.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            ngayKhenThuong = dtTemp.Rows[i][9].ToString();
                            if (ngayKhenThuong == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " năm xét khen thưởng không được để trống.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            else
                            {
                                if (kiemtrangaythangnam(ngayKhenThuong) == false || ngayKhenThuong.Length != 4)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " năm xet khen thưởng (" + dtTemp.Rows[i][9].ToString() + ") không đúng định dạng yyyy.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {

                                    try
                                    {

                                        ngayKhenThuong = dtTemp.Rows[i][9].ToString();
                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có năm xét khen thưởng (" + dtTemp.Rows[i][9].ToString() + ") không đúng định dạng yyyy.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }
                            ngayKyQD = dtTemp.Rows[i][8].ToString();
                            if (ngayKyQD == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " ngày ký quyết định không được để trống.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            else
                            {
                                if (kiemtrangaythangnam(ngayKyQD) == false)
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " ngày ký quyết định (" + dtTemp.Rows[i][8].ToString() + ") không đúng định dạng dd/mm/yyyy.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                                else
                                {

                                    try
                                    {

                                        ngayKyQD = DateTime.Parse(dtTemp.Rows[i][8].ToString().Trim().Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có ngày ký quyết định (" + dtTemp.Rows[i][8].ToString() + ") không đúng định dạng dd/mm/yyyy.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                            }
                            DataTable tblDanhHieu_HinhThuc = new DataTable();
                            tblDanhHieu_HinhThuc = _class.GetData("SELECT top 1 maKhenThuongpr,capDo,laThiDua FROM dbo.tblDMKhenThuong WHERE loaiDoiTuong LIKE N'%Cá nhân%' and tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");

                            if (tblDanhHieu_HinhThuc.Rows.Count > 0)//làm tới đây
                            {
                                danhHieu = tblDanhHieu_HinhThuc.Rows[0]["maKhenThuongpr"].ToString();
                            }
                            else
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " danh hiệu '" + dtTemp.Rows[i][10] + "' không tồn tại trong dữ liệu.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }

                            string sql_danhHieu_hinhThuc = _class.GetOneStringField("SELECT maKhenThuongpr FROM dbo.tblDMKhenThuong WHERE loaiDoiTuong LIKE N'%Cá nhân%' and tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");

                            hinhThuc = _class.GetOneStringField("SELECT maHinhThucTDKTpr FROM dbo.tblDMHinhThucTDKT WHERE tenHinhThucTDKT =N'" + dtTemp.Rows[i][11] + "'");
                            if (hinhThuc == "")
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có hình thức khen thưởng '" + dtTemp.Rows[i][11] + "' không tồn tại trong dữ liệu.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            capQuyetDinh = _class.GetOneStringField("SELECT top 1 convert(nvarchar(50),capDo) FROM dbo.tblDMKhenThuong WHERE tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");
                            loaiDoiTuong = "Cá nhân";
                            nhapExcel = "1";
                            string sttdoituongtapthe = "";
                            string sttDKTD = "";
                            DataTable tableDoiTuong = new DataTable();
                            DataTable tableDKTD = new DataTable();
                            //kiểm tra đối tượng
                            string check_doituong = "SELECT top 1 sttDoiTuongTDKTpr,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE tenDoiTuongTDKT = N'" + hoTen + "' and loaiDoiTuong =N'Cá nhân' " + (namSinh == "" ? "" : "and ngaySinh ='" + namSinh + "' ") + " and maDonVipr_noicongtac = '" + donVi + "'";
                            tableDoiTuong = _class.GetData(check_doituong);
                            if (tableDoiTuong.Rows.Count > 0)
                            {
                                sttDoiTuong = tableDoiTuong.Rows[0]["sttDoiTuongTDKTpr"].ToString();
                                //kiểm tra có tồn tại đăng ký hay không
                                string check_DKTD = "SELECT top 1 sttTDKTpr FROM dbo.tblThiDuaKhenThuong WHERE sttDoiTuongTDKTpr_sd =N'" + sttDoiTuong + "' AND maKhenThuongpr_sd =N'" + danhHieu + "' AND maHinhThucTDKTpr_sd =N'" + hinhThuc + "' AND sttDKThiDuapr_sd is not null AND (ngayDangKy between '" + HttpContext.Current.Session.GetNgayDauKy() + @"' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + @"')";
                                tableDKTD = _class.GetData(check_DKTD);
                                if (tableDKTD.Rows.Count > 0)
                                {
                                    sttDKTD = tableDKTD.Rows[0]["sttTDKTpr"].ToString();
                                    //kiểm tra điều kiện để nhập excel
                                    //string ngayDauKy = HttpContext.Current.Session.GetNgayDauKy().Substring(3, 2) + "/" + HttpContext.Current.Session.GetNgayDauKy().Substring(0, 2) + "/" + HttpContext.Current.Session.GetNgayDauKy().Substring(6, 4);
                                    //string ngayCuoiKy = HttpContext.Current.Session.GetNgayCuoiKy().Substring(3, 2) + "/" + HttpContext.Current.Session.GetNgayCuoiKy().Substring(0, 2) + "/" + HttpContext.Current.Session.GetNgayCuoiKy().Substring(6, 4);
                                    //decimal slSKKN = ntsLibraryFunctions.kiemTraSKKN(danhHieu.ToString(), sttDoiTuong.ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                                    //decimal slDHTD = ntsLibraryFunctions.kiemTraDKBatBuoc(danhHieu.ToString(), sttDoiTuong.ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                                    //decimal slDHTDPhu = ntsLibraryFunctions.kiemTraDKPhu(danhHieu.ToString(), sttDoiTuong.ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                                    //nếu trong khoảng thời gian có bị kỹ luật sẽ không được xét
                                    if (ntsLibraryFunctions.kiemTraDieuKienKhongXetTDKT(sttDoiTuong.ToString(), danhHieu.ToString()) == true)
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có danh hiệu '" + dtTemp.Rows[i][10] + "' không đủ điều kiện để nhập khen thưởng (Không xét thi đua khen thưởng)</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                    string ngayDauKy = HttpContext.Current.Session.GetNgayDauKy().Substring(3, 2) + "/" + HttpContext.Current.Session.GetNgayDauKy().Substring(0, 2) + "/" + HttpContext.Current.Session.GetNgayDauKy().Substring(6, 4);
                                    string ngayCuoiKy = HttpContext.Current.Session.GetNgayCuoiKy().Substring(3, 2) + "/" + HttpContext.Current.Session.GetNgayCuoiKy().Substring(0, 2) + "/" + HttpContext.Current.Session.GetNgayCuoiKy().Substring(6, 4);
                                    decimal slSKKN = ntsLibraryFunctions.kiemTraSKKN(danhHieu.ToString(), sttDoiTuong.ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                                    decimal slDHTD = ntsLibraryFunctions.kiemTraDKBatBuoc(danhHieu.ToString(), sttDoiTuong.ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);

                                    decimal slDHTDPhu = ntsLibraryFunctions.kiemTraDKPhu(danhHieu.ToString(), sttDoiTuong.ToString(), Convert.ToDateTime(ngayDauKy), Convert.ToDateTime(ngayCuoiKy), false);
                                    //if ((slSKKN + slDHTD + slDHTDPhu) >= 2)
                                    //{
                                    if ((slSKKN + slDHTD + slDHTDPhu) < 2)
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có danh hiệu '" + dtTemp.Rows[i][10] + "' không đủ điều kiện để nhập khen thưởng.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                                else
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " có danh hiệu '" + dtTemp.Rows[i][10] + "' chưa đăng ký thi đua.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }


                                //kiểm tra đối tượng tập thể
                                DataTable tableDoiTuong_TapThe = new DataTable();
                                string check_doituong_tapthe = "SELECT top 1 sttDoiTuongTDKTpr,tenChucVuQL,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE  tenDoiTuongTDKT = N'" + dtTemp.Rows[i][5] + "' AND maDonVipr_noicongtac = '" + donVi + "' and loaiDoiTuong =N'Tập thể'";
                                tableDoiTuong_TapThe = _class.GetData(check_doituong_tapthe);
                                if (tableDoiTuong_TapThe.Rows.Count > 0)
                                {
                                    sttdoituongtapthe = tableDoiTuong_TapThe.Rows[0]["sttDoiTuongTDKTpr"].ToString();
                                }
                                else
                                {


                                    sttdoituongtapthe = "";

                                }
                            }
                            else
                            {
                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " không tồn tại trong hệ thống danh mục cá nhân.</br>";
                                sodongthemkhongthanhcong += 1;
                                continue;
                            }
                            //nếu trong cấu hình cho phép nhập nhiều lần thì kiểm tra hoặc trong cấu hình chưa có thì kiểm tra
                            if (_class.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + danhHieu + "' ") == "0" || _class.GetOneStringField(@"select convert(nvarchar(2),nhapHTKTNhieuLan) from tblCauHinhTDKT where maKhenThuongpr_sd =N'" + danhHieu + "' ") == "")
                            {
                                //kiểm tra đối tượng đã được nhập khen thưởng hay chưa
                                if (_class.CheckHasRecord(@"select sttDoiTuongTDKTpr_sd from tblThiDuaKhenThuong where sttDoiTuongTDKTpr_sd ='" + sttDoiTuong + "' and ngayKyQD is not null and maKhenThuongpr_sd = '" + danhHieu + "' and maHinhThucTDKTpr_sd = '" + hinhThuc + "' AND sttDKThiDuapr_sd is not null  AND (ngayDangKy between '" + HttpContext.Current.Session.GetNgayDauKy() + @"' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + @"')") == false)
                                {
                                    //kiểm tra quyết định
                                    DataTable tableQuyetDinh = new DataTable();

                                    string stringQuyetDinh = "SELECT sttQuyetDinhKTpr FROM dbo.tblQuyetDinhKT WHERE maDonVipr_sd =N'" + HttpContext.Current.Session.GetMaDonVi() + "' AND maKhenThuongpr_sd =N'" + danhHieu + "' AND maHinhThucTDKTpr_sd =N'" + hinhThuc + "' AND soQuyetDinh=N'" + soQuyetDinh + "' AND ngayKy=N'" + ngayKyQD + "' AND capQuyetDinh=N'" + capQuyetDinh + "' AND loaiDoiTuong=N'" + loaiDoiTuong + "'";
                                    tableQuyetDinh = _class.GetData(stringQuyetDinh);
                                    if (tableQuyetDinh.Rows.Count > 0)
                                    {
                                        sttQD = tableQuyetDinh.Rows[0]["sttQuyetDinhKTpr"].ToString();
                                    }
                                    else
                                    {
                                        try
                                        {
                                            string sqlInsert_ = @"INSERT INTO [dbo].[tblQuyetDinhKT] ([ngayKhenThuong], [maKhenThuongpr_sd], [maHinhThucTDKTpr_sd], [soQuyetDinh], [ngayKy], [capQuyetDinh], [ngayThaoTac], [nguoiThaoTac], [maDonVipr_sd],[loaiDoiTuong],[nhapExcel],namXetKhenThuong) values
	                            ( N'" + ngayKyQD + "', N'" + danhHieu + "', N'" + hinhThuc + "', N'" + soQuyetDinh + "',N'" + ngayKyQD + "',N'" + capQuyetDinh + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + loaiDoiTuong + "','1', N'" + ngayKhenThuong.ToString() + "')";
                                            if (_class.ExeCuteNonQuery(sqlInsert_) == true)
                                            {
                                                string sqlchuoi = "SELECT CONVERT(DECIMAL,MAX(sttQuyetDinhKTpr)) FROM dbo.tblQuyetDinhKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                                decimal sttQD_ = _class.GetOneDecimalField(sqlchuoi);

                                                sttQD = sttQD_.ToString();
                                            }
                                            else
                                            {
                                                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                                sodongthemkhongthanhcong += 1;
                                                continue;
                                            }
                                        }
                                        catch
                                        {
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                            sodongthemkhongthanhcong += 1;
                                            continue;
                                        }
                                    }
                                    try
                                    {
                                        //update đối tượng vào quyết định :tblthiduakhenthuong

                                        string sqlInsertkt = @"UPDATE [dbo].[tblThiDuaKhenThuong] set  tinhTrang =N'Có quyết định',soQuyetDinh =  N'" + soQuyetDinh + "',ngayKhenThuong=N'" + ngayKyQD + "',ngayKyQD=N'" + ngayKyQD + "',capQuyetDinh=N'" + capQuyetDinh + "',sttQuyetDinhKTpr_sd=N'" + sttQD + "',nhapExcel='1',namXetKhenThuong = N'" + ngayKhenThuong.ToString() + "' where sttTDKTpr ='" + sttDKTD + "'";
                                        _class.ExeCuteNonQuery(sqlInsertkt);
                                        sodongthemthanhcong = sodongthemthanhcong + 1;


                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "-Cập nhật quyết định '" + soQuyetDinh + "' khen thưởng cho cá nhân: " + hoTen + " không thành công.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }

                                }
                                else
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Cá nhân: " + hoTen + " đã được nhập khen thưởng số '" + soQuyetDinh + "' ngày ký quyết định '" + dtTemp.Rows[i][8].ToString() + "' rồi.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                            }
                            else
                            {
                                //kiểm tra quyết định
                                DataTable tableQuyetDinh = new DataTable();

                                string stringQuyetDinh = "SELECT sttQuyetDinhKTpr FROM dbo.tblQuyetDinhKT WHERE maDonVipr_sd =N'" + HttpContext.Current.Session.GetMaDonVi() + "'  AND maKhenThuongpr_sd =N'" + danhHieu + "' AND maHinhThucTDKTpr_sd =N'" + hinhThuc + "' AND soQuyetDinh=N'" + soQuyetDinh + "' AND ngayKy=N'" + ngayKyQD + "' AND capQuyetDinh=N'" + capQuyetDinh + "' AND loaiDoiTuong=N'" + loaiDoiTuong + "'";
                                tableQuyetDinh = _class.GetData(stringQuyetDinh);
                                if (tableQuyetDinh.Rows.Count > 0)
                                {
                                    sttQD = tableQuyetDinh.Rows[0]["sttQuyetDinhKTpr"].ToString();
                                }
                                else
                                {
                                    try
                                    {
                                        string sqlInsert_ = @"INSERT INTO [dbo].[tblQuyetDinhKT] ([ngayKhenThuong], [maKhenThuongpr_sd], [maHinhThucTDKTpr_sd], [soQuyetDinh], [ngayKy], [capQuyetDinh], [ngayThaoTac], [nguoiThaoTac], [maDonVipr_sd],[loaiDoiTuong],[nhapExcel],namXetKhenThuong) values
	                            ( N'" + ngayKyQD + "', N'" + danhHieu + "', N'" + hinhThuc + "', N'" + soQuyetDinh + "',N'" + ngayKyQD + "',N'" + capQuyetDinh + "',N'" + HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy") + "',N'" + HttpContext.Current.Session.GetCurrentUserID() + "',N'" + HttpContext.Current.Session.GetMaDonVi() + "',N'" + loaiDoiTuong + "','1', N'" + ngayKhenThuong.ToString() + "')";
                                        if (_class.ExeCuteNonQuery(sqlInsert_) == true)
                                        {
                                            string sqlchuoi = "SELECT CONVERT(DECIMAL,MAX(sttQuyetDinhKTpr)) FROM dbo.tblQuyetDinhKT WHERE maDonvipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'";
                                            decimal sttQD_ = _class.GetOneDecimalField(sqlchuoi);

                                            sttQD = sttQD_.ToString();
                                        }
                                        else
                                        {
                                            ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                            sodongthemkhongthanhcong += 1;
                                            continue;
                                        }
                                    }
                                    catch
                                    {
                                        ssloinhapexcelcanhan = ssloinhapexcelcanhan + "- Thêm mới quyết định '" + soQuyetDinh + "' cho cá nhân: " + hoTen + " không thành công.</br>";
                                        sodongthemkhongthanhcong += 1;
                                        continue;
                                    }
                                }
                                try
                                {
                                    //update đối tượng vào quyết định :tblthiduakhenthuong

                                    string sqlInsertkt = @"UPDATE [dbo].[tblThiDuaKhenThuong] set  tinhTrang =N'Có quyết định',soQuyetDinh =  N'" + soQuyetDinh + "',ngayKhenThuong=N'" + ngayKyQD + "',ngayKyQD=N'" + ngayKyQD + "',capQuyetDinh=N'" + capQuyetDinh + "',sttQuyetDinhKTpr_sd=N'" + sttQD + "',nhapExcel='1',namXetKhenThuong = N'" + ngayKhenThuong.ToString() + "' where sttTDKTpr ='" + sttDKTD + "'";
                                    _class.ExeCuteNonQuery(sqlInsertkt);
                                    sodongthemthanhcong = sodongthemthanhcong + 1;


                                }
                                catch
                                {
                                    ssloinhapexcelcanhan = ssloinhapexcelcanhan + "-Cập nhật quyết định '" + soQuyetDinh + "' khen thưởng cho cá nhân: " + hoTen + " không thành công.</br>";
                                    sodongthemkhongthanhcong += 1;
                                    continue;
                                }
                            }

                        }
                        else
                        { break; }
                    }
                }

                Session["ssloinhapexcelcanhan"] = ssloinhapexcelcanhan;

                if (sodongthemkhongthanhcong == 0)
                {
                    return "Đã nhận: " + sodongthemthanhcong.ToString() + "/" + sodongtong.ToString() + " dòng.";
                }
                else
                {
                    return "Đã nhận: " + sodongthemthanhcong.ToString() + "/" + sodongtong.ToString() + " dòng. Có: " + sodongthemkhongthanhcong.ToString() + " dòng bị sai thông tin:<a href='#' onclick='hienThongBaoLoi();return false;'> Bấm xem chi tiết</a>";
                }
            }
            catch (Exception ex)
            {

                ssloinhapexcelcanhan = ssloinhapexcelcanhan + "Tải tập tin excel không thành công. lý do " + ex.Message.ToString() + "</br>";
                return ssloinhapexcelcanhan;
            }
        }

        [AjaxPro.AjaxMethod]
        public string soLuongTonTai()
        {
            string path = "";

            decimal sodongquyetdinh = 0;
            decimal sodongphongban = 0;
            decimal sodongcanhan = 0;
            decimal sodongtong = 0;
            string noidungqd = "";
            string sttqdtam = "";
            if (Session["ssGetUrlFile"] != null)
            {
                path = Session["ssGetUrlFile"].ToString();
            }
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
                string sqlSelect = "select * from [KhenThuong$A8:L20000]";
                OleDbCommand conmand = new OleDbCommand(sqlSelect, connection);
                OleDbDataAdapter oleDA = new OleDbDataAdapter(conmand);
                oleDA.Fill(dtTemp);
                connection.Close();
                string noidungcanhan = "";
                //sodongtong = dtTemp.Rows.Count.ToString();
                int sodong = 0, n = 0;
                string sqlInsert = "", sttDoiTuong = "", sttQD = "", hoTen = "", chucVu = "", namSinh = "", donVi = "", noiCongTac = "", soQuyetDinh = "", ngayKhenThuong = "", ngayKyQD = "", danhHieu = "", hinhThuc = "", capQuyetDinh = "", loaiDoiTuong = "", nhapExcel = "";
                string maDV = HttpContext.Current.Session.GetMaDonVi();
                DataTable tblDonViTrucThuoc = new DataTable();
                string strWhereDV = "maDonVipr IN (SELECT maDonvipr FROM dbo.tblDMDonvi WHERE (maDonvipr= N'" + maDV + "' OR (maDonvipr_cha= N'" + maDV + "') OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "')  OR maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "'))) OR  maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha IN (SELECT maDonvipr FROM tblDMDonvi WHERE maDonvipr_cha= N'" + maDV + "') ) ) )";
                string sqlDonViTrucThuoc = " select maDonVipr,tenDonVi from tbldmDonVi where " + strWhereDV;
                tblDonViTrucThuoc = _class.GetData(sqlDonViTrucThuoc);
                decimal soluongdanhxung = 0;
                string donvitontai = "";
                if (dtTemp.Columns.Count > 0)
                {
                    for (int i = 0; i < dtTemp.Rows.Count; i++)
                    {
                        if (dtTemp.Rows[i][2].ToString() != "")
                        {
                            sodongtong = sodongtong + 1;
                            hoTen = dtTemp.Rows[i][2].ToString();
                            bool loi = false;
                            bool loi_danhxung = false;
                            string[] danhXung = { "Ông", "Bà", "ông", "bà", "ba", "Ong", "Ba", "ong", "Mrs", "Mr", "mrs", "mr" };
                            for (int y = 0; y < danhXung.Length; y++)
                            {
                                if (hoTen.IndexOf(danhXung[y].ToString()) == 0)
                                {
                                    loi_danhxung = true;
                                    break;
                                }
                            }
                            if (loi_danhxung)
                            {
                                soluongdanhxung = soluongdanhxung + 1;
                                noidungcanhan = noidungcanhan + ", " + hoTen;
                                continue;
                            }
                            string[] dauDacBiet = { ".", "@", "#", "$", "%", "^", "&", "*", "=" };
                            for (int y = 0; y < dauDacBiet.Length; y++)
                            {

                                if (hoTen.IndexOf(dauDacBiet[y].ToString()) >= 0)
                                {
                                    loi = true;
                                    break;
                                }
                            }
                            if (loi)
                            {
                                continue;
                            }

                            chucVu = dtTemp.Rows[i][3].ToString();
                            namSinh = dtTemp.Rows[i][4].ToString();
                            if (namSinh != "")
                            {
                                if (kiemtrangaythangnam(namSinh) == false)
                                {
                                    continue;

                                }
                                else
                                {
                                    if (namSinh.Length == 4)
                                    {
                                        string ngay = "01/01/" + namSinh.ToString().Trim();
                                        namSinh = DateTime.Parse(ngay.Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                    }
                                    else
                                    {
                                        try
                                        {
                                            namSinh = DateTime.Parse(namSinh.ToString().Trim().Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                                        }
                                        catch
                                        {
                                            continue;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                namSinh = "";
                            }
                            if (dtTemp.Rows[i][6].ToString() == "")
                            {
                                continue;
                            }
                            else
                            {
                                decimal soluongdonvi = _class.GetOneDecimalField(@"select convert(decimal(18,0),count(maDonvipr)) from tbldmDonVi where tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "' and " + strWhereDV + "");
                                if (soluongdonvi > 1)
                                {
                                    donvitontai = "Trong hệ thống tồn tại nhiều hơn 1 đơn vị trực thuộc";
                                    break;
                                }
                                donVi = _class.GetOneStringField("SELECT top 1 maDonvipr FROM dbo.tblDMDonvi WHERE tenDonVi = N'" + dtTemp.Rows[i][6].ToString().Trim() + "' and " + strWhereDV + "").ToString();
                                if (donVi.ToString() == "")
                                {
                                    continue;
                                }
                                else
                                {
                                    string timthay = "";
                                    for (int j = 0; j < tblDonViTrucThuoc.Rows.Count; j++)
                                    {
                                        if (donVi.ToString() == tblDonViTrucThuoc.Rows[j]["maDonvipr"].ToString())
                                        {
                                            timthay = "timthay";
                                            break;
                                        }

                                    }
                                    if (timthay.ToString() != "timthay")
                                    {
                                        continue;
                                    }
                                }
                            }
                            soQuyetDinh = dtTemp.Rows[i][7].ToString();
                            if (soQuyetDinh == "")
                            {
                                continue;
                            }
                            ngayKhenThuong = dtTemp.Rows[i][9].ToString();
                            if (ngayKhenThuong == "")
                            {
                                continue;
                            }
                            try
                            {
                                ngayKhenThuong = dtTemp.Rows[i][9].ToString();
                            }
                            catch
                            {
                                continue;
                            }
                            ngayKyQD = dtTemp.Rows[i][8].ToString();
                            if (ngayKyQD == "")
                            {
                                continue;
                            }
                            try
                            {
                                ngayKyQD = DateTime.Parse(dtTemp.Rows[i][8].ToString().Trim().Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("MM/dd/yyyy");
                            }
                            catch
                            {
                                continue;
                            }

                            danhHieu = _class.GetOneStringField("SELECT maKhenThuongpr FROM dbo.tblDMKhenThuong WHERE loaiDoiTuong LIKE N'%Cá nhân%' and tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");
                            if (danhHieu == "")
                            {

                                continue;
                            }

                            hinhThuc = _class.GetOneStringField("SELECT maHinhThucTDKTpr FROM dbo.tblDMHinhThucTDKT WHERE tenHinhThucTDKT =N'" + dtTemp.Rows[i][11] + "'");
                            if (hinhThuc == "")
                            {

                                continue;
                            }
                            capQuyetDinh = _class.GetOneStringField("SELECT top 1 convert(nvarchar(50),capDo) FROM dbo.tblDMKhenThuong WHERE tenKhenThuong =N'" + dtTemp.Rows[i][10] + "'");
                            loaiDoiTuong = "Cá nhân";
                            nhapExcel = "1";
                            string sttdoituongtapthe = "";
                            DataTable tableDoiTuong = new DataTable();
                            //kiểm tra đối tượng
                            string check_doituong = "SELECT top 1 sttDoiTuongTDKTpr,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE tenDoiTuongTDKT = N'" + hoTen + "' and loaiDoiTuong =N'Cá nhân' " + (namSinh == "" ? "" : "and ngaySinh ='" + namSinh + "' ") + " and maDonVipr_noicongtac = '" + donVi + "'";
                            tableDoiTuong = _class.GetData(check_doituong);
                            if (tableDoiTuong.Rows.Count > 0)
                            {
                                sttDoiTuong = tableDoiTuong.Rows[0]["sttDoiTuongTDKTpr"].ToString();

                                //kiểm tra đối tượng tập thể
                                DataTable tableDoiTuong_TapThe = new DataTable();
                                string check_doituong_tapthe = "SELECT top 1 sttDoiTuongTDKTpr,tenChucVuQL,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE  tenDoiTuongTDKT = N'" + dtTemp.Rows[i][5] + "' AND maDonVipr_noicongtac = '" + donVi + "' and loaiDoiTuong =N'Tập thể'";
                                tableDoiTuong_TapThe = _class.GetData(check_doituong_tapthe);
                                if (tableDoiTuong_TapThe.Rows.Count > 0)
                                {
                                    sttdoituongtapthe = tableDoiTuong_TapThe.Rows[0]["sttDoiTuongTDKTpr"].ToString();
                                }
                                else
                                {
                                    if (dtTemp.Rows[i][5].ToString() != "")
                                    {
                                        sodongphongban = sodongphongban + 1;
                                    }
                                }
                            }
                            else
                            {
                                DataTable tableDoiTuong_TapThe = new DataTable();
                                //kiểm tra đối tượng tập thể

                                string check_doituong_tapthe = "SELECT top 1 sttDoiTuongTDKTpr,tenChucVuQL,tenDoiTuongTDKT,ngaySinh,maDonVipr_noicongtac FROM dbo.tblDMDoiTuongTDKT WHERE  tenDoiTuongTDKT = N'" + dtTemp.Rows[i][5] + "' AND maDonVipr_noicongtac = '" + donVi + "' and loaiDoiTuong =N'Tập thể'";
                                tableDoiTuong_TapThe = _class.GetData(check_doituong_tapthe);
                                if (tableDoiTuong_TapThe.Rows.Count > 0)
                                {
                                    sttdoituongtapthe = tableDoiTuong_TapThe.Rows[0]["sttDoiTuongTDKTpr"].ToString();
                                }
                                else
                                {
                                    if (dtTemp.Rows[i][5].ToString() != "")
                                    {
                                        sodongphongban = sodongphongban + 1;
                                    }

                                }
                                sodongcanhan = sodongcanhan + 1;
                            }
                            //kiểm tra đối tượng đã được nhập khen thưởng hay chưa

                            //kiểm tra quyết định
                            DataTable tableQuyetDinh = new DataTable();

                            string stringQuyetDinh = "SELECT sttQuyetDinhKTpr, soQuyetDinh,CONVERT(NVARCHAR(10),ngayKy,103) AS ngayKy ,ngayKhenThuong,tenKhenThuong =(SELECT tenKhenThuong FROM dbo.tblDMKhenThuong WHERE maKhenThuongpr = maKhenThuongpr_sd) FROM dbo.tblQuyetDinhKT WHERE maDonVipr_sd =N'" + HttpContext.Current.Session.GetMaDonVi() + "'  AND maKhenThuongpr_sd =N'" + danhHieu + "' AND maHinhThucTDKTpr_sd =N'" + hinhThuc + "' AND soQuyetDinh=N'" + soQuyetDinh + "' AND ngayKy=N'" + ngayKyQD + "' AND capQuyetDinh=N'" + capQuyetDinh + "' AND loaiDoiTuong=N'" + loaiDoiTuong + "'";
                            tableQuyetDinh = _class.GetData(stringQuyetDinh);
                            if (tableQuyetDinh.Rows.Count > 0)
                            {
                                sttQD = tableQuyetDinh.Rows[0]["sttQuyetDinhKTpr"].ToString();
                                if (sttqdtam.ToString() != sttQD)
                                {
                                    sodongquyetdinh = sodongquyetdinh + 1;
                                    noidungqd = noidungqd + "+ Số quyết định: " + tableQuyetDinh.Rows[0]["soQuyetDinh"].ToString() + " có ngày ký: " + tableQuyetDinh.Rows[0]["ngayKy"].ToString() + " cho danh hiệu: " + tableQuyetDinh.Rows[0]["tenKhenThuong"].ToString() + " <br>";
                                }
                                sttqdtam = sttQD;

                            }


                        }
                        else
                        { break; }
                    }
                }
                if (donvitontai != "")
                {
                    return donvitontai;
                }
                if (sodongphongban == 0 && sodongcanhan == 0 && sodongquyetdinh == 0)
                {
                    return "khongco";
                }
                string kq = "";
                if (sodongphongban != 0)
                {
                    kq = kq + "- Có: " + sodongphongban.ToString() + " phòng ban sẽ thêm mới.<br> ";
                }
                if (sodongcanhan != 0)
                {
                    kq = kq + "- Có: " + sodongcanhan.ToString() + " cá nhân sẽ thêm mới. </br>";
                }
                if (sodongquyetdinh != 0)
                {
                    kq = kq + "- Có: " + sodongquyetdinh.ToString() + " số quyết định đã tồn tại trong hệ thống: " + noidungqd + " <br>";
                }
                if (soluongdanhxung != 0)
                {
                    kq = kq + "- Tồn tại: " + soluongdanhxung.ToString() + " cá nhân có danh xưng trong họ tên là: " + noidungcanhan + " <br>";
                }
                return kq + "Nếu đồng ý thực hiện hệ thống sẽ tự động thêm phòng ban hoặc cá nhân vào hệ thống đồng thời thêm tiếp cá nhân cho quyết định đã tồn tại. Bạn có đồng ý thực hiện?";
            }
            catch (Exception ex)
            {
                return "Nhận file excel thất bại! Lý do: " + ex.Message.ToString() + "";// daNhan.ToString() + "/" + soLuong.ToString();
            }

        }
        private bool kiemtrangaythangnam(string chuoikiemtra)
        {
            try
            {
                if (chuoikiemtra.Length > 10)
                {
                    chuoikiemtra = DateTime.Parse(chuoikiemtra.Substring(0, 10).ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb")).ToString("dd/MM/yyyy");
                }
                if (chuoikiemtra.Length == 10)
                {
                    string ngay = chuoikiemtra.ToString().Trim().Split('/')[0].ToString();
                    string thang = chuoikiemtra.ToString().Trim().Split('/')[1].ToString();
                    string nam = chuoikiemtra.ToString().Trim().Split('/')[2].ToString();
                    if (ngay.Length.ToString() != "2" || thang.Length.ToString() != "2" || nam.Length.ToString() != "4")
                    {
                        return false;
                    }
                    else
                    {
                        if (Convert.ToDecimal(ngay.ToString()) > 32 || Convert.ToDecimal(ngay.ToString()) < 0)
                        {
                            return false;
                        }
                        if (Convert.ToDecimal(thang.ToString()) > 12 || Convert.ToDecimal(thang.ToString()) < 0)
                        {
                            return false;
                        }
                        if (Convert.ToDecimal(nam.ToString()) > 2090 || Convert.ToDecimal(nam.ToString()) < 1900)
                        {
                            return false;
                        }
                    }

                }
                else
                {
                    if (chuoikiemtra.Length == 4)
                    {
                        if (Convert.ToDecimal(chuoikiemtra.ToString()) > 2090 || Convert.ToDecimal(chuoikiemtra.ToString()) < 1900)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                return true;

            }
            catch
            {
                return false;
            }
        }

        [AjaxPro.AjaxMethod]
        public string laydanhsachdoituong()
        {
            return Session["ssloinhapexcelcanhan"].ToString();
        }
        [AjaxPro.AjaxMethod]
        public string sodongthemkhongthanhcong()
        {
            return Session["sodongthemkhongthanhcong"].ToString();
        }
        [AjaxPro.AjaxMethod]
        public DataTable loadComBoxPhongBan(string maDonVipr_sd)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string sql = "select sttDoiTuongTDKTpr,tenDoiTuongTDKT from tblDMDoiTuongTDKT where loaiDoiTuong like N'%Tập thể%' and maDonVipr_noicongtac =N'" + maDonVipr_sd + "'";
            return sqlFun.GetData(sql);
        }
        [AjaxPro.AjaxMethod]
        public DataTable loadComBoxHoiDong(string namXetKhenThuong)
        {
            SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string sql = "select convert(varchar(18),sttHoiDongXDTDpr) as sttHoiDongXDTDpr,tenHoiDong from tblHoiDongXetDuyetTD where maDonVipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "' and namXetDuyet=N'" + namXetKhenThuong.ToString() + "' and ngungTD = 0";
            return sqlFun.GetData(sql);
        }
        [AjaxPro.AjaxMethod]
        public DataTable loadcboHDSKKN(string ngayduyet)
        {
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string sql = @"SELECT sttHoiDongXDSKKNpr,tenHoiDong FROM dbo.tblHoiDongXetDuyetSKKN 
                WHERE  '" + _mChuyenChuoiSangNgay(ngayduyet).Substring(6, 4) + "' = YEAR(ngayThanhLap) AND maDonVipr_sd =N'" + HttpContext.Current.Session.GetMaDonVi() + "' ";
                return sqlFun.GetData(sql);
            }
            catch
            {
                return null;
            }
        }
        [AjaxPro.AjaxMethod]
        public string getNamThaoTac()
        {
            return HttpContext.Current.Session.getNamSudung();
        }
        // hoi dong skkn
        [AjaxPro.AjaxMethod]
        public string insertGrid1(AjaxPro.JavaScriptArray param)
        {
            try
            {
                string sql = @"INSERT INTO dbo.tblHoiDongXetDuyetTD (tenHoiDong,ngayThanhLap,namXetDuyet,ghiChu,ngungTD,maDonVipr_sd,nguoiThaoTac,ngayThaotac)
                 VALUES(@tenHoiDong,@ngayThanhLap,@namXetDuyet,@ghiChu,@ngungTD,@maDonVipr_sd,@nguoiThaoTac,@ngayThaotac)";
                SqlConnection sqlCon = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
                SqlFunction sqlFunc = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                sqlCon.Open();
                SqlCommand cmd = new SqlCommand(sql, sqlCon);
                cmd.Parameters.Add(new SqlParameter("@tenHoiDong", param[0].ToString()));
                cmd.Parameters.Add(new SqlParameter("@ngayThanhLap", _mChuyenChuoiSangNgay(param[1].ToString())));
                cmd.Parameters.Add(new SqlParameter("@namXetDuyet", param[2].ToString()));
                cmd.Parameters.Add(new SqlParameter("@ghiChu", param[3].ToString()));
                cmd.Parameters.Add(new SqlParameter("@ngungTD", "0"));
                cmd.Parameters.Add(new SqlParameter("@maDonVipr_sd", HttpContext.Current.Session.GetMaDonVi()));
                cmd.Parameters.Add(new SqlParameter("@nguoiThaoTac", HttpContext.Current.Session.GetCurrentUserID()));
                cmd.Parameters.Add(new SqlParameter("@ngayThaotac", HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy")));
                cmd.ExecuteNonQuery();
                sqlCon.Close();
                cmd.Dispose();
                return sqlFunc.GetOneDecimalField("SELECT CONVERT(DECIMAL,MAX(sttHoiDongXDTDpr)) FROM dbo.tblHoiDongXetDuyetTD WHERE maDonVipr_sd=N'" + HttpContext.Current.Session.GetMaDonVi() + "'").ToString();
            }
            catch
            {
                return "0";
            }
        }

        [AjaxPro.AjaxMethod]
        public string kiemTraTaiExcelMau()
        {

            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());

                //                string sql = @"SELECT tenDoiTuongTDKT,tenChucVuQL
                //                            , ngaySinh = (case when (namSinh ='1' OR namSinh IS NULL) then convert(nvarchar(4),year(ngaySinh)) else convert(nvarchar(10),ngaysinh,103) end)
                //                            ,(SELECT tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT a WHERE a.sttDoiTuongTDKTpr = tdkt.sttDoiTuongTDKTpr_cha) AS noiCongTac
                //                            ,(SELECT tenDonVi FROM dbo.tblDMDonvi WHERE maDonvipr = maDonVipr_noicongtac) AS donVi
                //                            FROM dbo.tblDMDoiTuongTDKT tdkt WHERE tdkt.sttDoiTuongTDKTpr_sd in (select sttDoiTuongTDKTpr_sd from tblDangKyThiDua where ngayDangKy between '" + HttpContext.Current.Session.GetNgayDauKy() + @"' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + @"') and  maDonVipr_noicongtac = N'" + HttpContext.Current.Session.GetMaDonVi() + "' AND loaiDoiTuong=N'Cá nhân'";

                string sql = @"SELECT 
                danhxung = (SELECT danhXungpr_sd FROM dbo.tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd),
                hoten = (SELECT tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd),
                chucvu = (SELECT tenChucVuQL FROM dbo.tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd),
                ngaysinh = (SELECT ngaySinh=(CASE when namSinh ='1'  then convert(nvarchar(4),year(ngaySinh)) else convert(nvarchar(10),ngaysinh,103) end) from tblDMDoiTuongTDKT where sttDoiTuongTDKTpr = sttDoiTuongTDKTpr_sd),
                noicongtac = (SELECT tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT b WHERE b.sttDoiTuongTDKTpr = isnull(tblThiDuaKhenThuong.sttDoiTuongTDKTpr_cha,0)),
                donvi = (SELECT (SELECT tenDonVi FROM dbo.tblDMDonvi WHERE maDonvipr=a.maDonVipr_noicongtac) FROM dbo.tblDMDoiTuongTDKT a WHERE sttDoiTuongTDKTpr=sttDoiTuongTDKTpr_sd),
                danhhieu = (SELECT tenKhenThuong FROM dbo.tblDMKhenThuong WHERE maKhenThuongpr = maKhenThuongpr_sd),
                hinhthuc = (SELECT tenHinhThucTDKT FROM dbo.tblDMHinhThucTDKT WHERE maHinhThucTDKTpr = maHinhThucTDKTpr_sd)
                FROM dbo.tblThiDuaKhenThuong WHERE sttDoiTuongTDKTpr_sd in (select sttDoiTuongTDKTpr FROM dbo.tblDMDoiTuongTDKT where maDonVipr_noicongtac = N'" + HttpContext.Current.Session.GetMaDonVi() + "') and  loaiDoiTuong=N'Cá nhân' AND (ngayDangKy between '" + HttpContext.Current.Session.GetNgayDauKy() + @"' and '" + HttpContext.Current.Session.GetNgayCuoiKy() + @"') order by sttDoiTuongTDKTpr_sd ";
                DataTable table_canhan = sqlFun.GetData(sql);
                if (table_canhan.Rows.Count > 0)
                {

                    int vDongXuat = 10;
                    string url = "~/xuatexcel" + "/" + HttpContext.Current.Session.GetCurrentUserID() + "/";
                    string fileName = "ExcelMauKQTDKTCaNhan" + HttpContext.Current.Session.GetMaDonVi() + ".xlsx";
                    if (!System.IO.Directory.Exists(Server.MapPath(url)))
                    {
                        System.IO.Directory.CreateDirectory(Server.MapPath(url));
                    }
                    DirectoryInfo di = new DirectoryInfo(Server.MapPath(url));
                    FileInfo[] rgFiles = di.GetFiles();
                    foreach (FileInfo fi in rgFiles)
                    {
                        fi.Delete();
                    }
                    File.Copy(Server.MapPath("~/excelmau/ExcelMauKQTDKTCaNhan.xlsx"), Server.MapPath("~/xuatexcel" + "/" + HttpContext.Current.Session.GetCurrentUserID() + "/" + fileName), true);
                    var wb = new XLWorkbook(Server.MapPath("~/xuatexcel" + "/" + HttpContext.Current.Session.GetCurrentUserID() + "/" + fileName));
                    var ws = wb.Worksheet("KhenThuong");
                    int stt = 0;
                    foreach (DataRow dr in table_canhan.Rows)
                    {
                        stt += 1;
                        ws.Range("A9:M9").CopyTo(ws.Range("A" + vDongXuat + ":L" + vDongXuat));
                        ws.Cell(vDongXuat, "A").Value = stt.ToString();
                        ws.Cell(vDongXuat, "B").Value = dr["danhxung"].ToString();
                        ws.Cell(vDongXuat, "C").Value = dr["hoten"].ToString();
                        ws.Cell(vDongXuat, "D").Value = dr["chucvu"].ToString();
                        ws.Cell(vDongXuat, "E").Value = "'" + dr["ngaySinh"].ToString();
                        ws.Cell(vDongXuat, "F").Value = dr["noiCongTac"].ToString();
                        ws.Cell(vDongXuat, "G").Value = dr["donVi"].ToString();
                        ws.Cell(vDongXuat, "H").Value = "";
                        ws.Cell(vDongXuat, "I").Value = "";
                        ws.Cell(vDongXuat, "J").Value = "";
                        ws.Cell(vDongXuat, "K").Value = dr["danhhieu"].ToString();
                        ws.Cell(vDongXuat, "L").Value = dr["hinhthuc"].ToString();
                        ws.Cell(vDongXuat, "M").Value = "";
                        vDongXuat += 1;
                    }
                    ws.Range("A9:M9").Delete(XLShiftDeletedCells.ShiftCellsUp);
                    wb.Save();
                    //
                    vDongXuat = 2;
                    ws = wb.Worksheet("DanhMuc");
                    DataTable table_phongban = sqlFun.GetData(@"SELECT tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT WHERE maDonVipr_noicongtac = N'" + HttpContext.Current.Session.GetMaDonVi() + "'  AND loaiDoiTuong=N'Tập thể'");
                    foreach (DataRow dr in table_phongban.Rows)
                    {
                        ws.Cell(vDongXuat, "D").Value = dr["tenDoiTuongTDKT"].ToString();
                        vDongXuat += 1;
                    }
                    wb.Save();
                    string hostName = HttpContext.Current.Request.Url.Host;
                    return "/xuatexcel/" + HttpContext.Current.Session.GetCurrentUserID() + "/" + fileName;

                }
                else
                {
                    return "/ExcelMau/ExcelMauKQTDKTCaNhan.xlsx";
                }
            }
            catch
            {
                return "/ExcelMau/ExcelMauKQTDKTCaNhan.xlsx";
            }
        }

        // hàm kiểm tra đối tượng đủ điều kiện đăng ký tđkt cho danh hiệu / hình thức hay không
        public bool ketQuaXetDieuKien(string maKhenThuong, string sttDoiTuongTDKT, string namXetDangKy, string dangKyExcel)
        {
            //lấy năm xét tiêu chuẩn là năm học hay năm hành chính
            SqlFunction class_ = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            decimal namhanhchinh;
            if (class_.CheckHasRecord(@"SELECT xetKhenThuongTheoNH FROM tblDMDonvi WHERE maDonVipr =N'" + HttpContext.Current.Session.GetMaDonVi() + "' and isnull(xetKhenThuongTheoNH,0) = '0' ") == true)
            {
                namhanhchinh = 0;
            }
            else
            {
                namhanhchinh = 1;
            }
            string tieuchuan = "dutieuchuan";
            try
            {
                SqlFunction sqlFun = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                string sql = "SELECT sttKhongXetTDKTpr FROM tblKhongXetTDKT WHERE sttDoiTuongTDKTpr_sd =N'" + sttDoiTuongTDKT + "' and namXetDuyet = N'" + namXetDangKy + "'";
                DataTable tblKhongXetTDKT = new DataTable();
                tblKhongXetTDKT = sqlFun.GetData(sql);
                //1. nếu tồn tại trong  bảng không xet thì duyệt tìm danh hiệu nếu có xuất ra không đủ điều kiện, không có thì mới tiếp bảng cấu hình
                if (tblKhongXetTDKT.Rows.Count > 0)
                {

                    for (int i = 0; i < tblKhongXetTDKT.Rows.Count; i++)
                    {
                        DataTable tblKhongXetTDKTCT = new DataTable();
                        tblKhongXetTDKTCT = sqlFun.GetData("SELECT maKhenThuongpr_sd = (select maKhenThuongpr_sd FROM tblDMKhongXetTDKT WHERE maKhongXetTDKTpr = maKhongXetTDKTpr_sd) FROM dbo.tblKhongXetTDKTCT WHERE sttKhongXetTDKTpr_sd =N'" + tblKhongXetTDKT.Rows[i][0].ToString() + "'");
                        for (int j = 0; j < tblKhongXetTDKTCT.Rows.Count; j++)
                        {
                            if (tblKhongXetTDKTCT.Rows[j][0].ToString() == maKhenThuong.ToString())
                            {
                                tieuchuan = "khongdutieuchuan";//đã tìm được 
                                return false;
                            }
                        }
                    }
                }
                //2nếu không có trong bảng không xét xét tiếp qua cấu hình nếu k có trong cấu hình thì đủ điều kiện
                DataTable tblCauHinh = new DataTable();
                tblCauHinh = sqlFun.GetData("SELECT sttCauHinhpr,coSKKN,SKKNCS,SKKNTinh,SKKNTQ,toanTu FROM tblCauHinhTDKT  WHERE maKhenThuongpr_sd = N'" + maKhenThuong + "' AND  (coSKKN <> '0' or coDKTD <> '0' OR sttCauHinhpr IN (SELECT sttCauHinhpr_sd FROM dbo.tblCauHinhTDKTCT))");
                if (tblCauHinh.Rows.Count > 0)//nếu có trong bảng cấu hình
                {
                    //1. kiểm nếu toán tử nếu hoặc thì chỉ cần một tiêu chí đủ đk thì đủ điều kiện

                    if (tblCauHinh.Rows[0]["toanTu"].ToString() == "Hoặc")
                    {
                        //nếu check có skkn được check và số lương skkncs hoặc skkntinh > 1 thì xét điều kiện
                        if (tblCauHinh.Rows[0]["coSKKN"].ToString() == "True")
                        {
                            //lấy số năm xét để xét skkn
                            DataTable tblCauHinhCT_SKKN = new DataTable();
                            tblCauHinhCT_SKKN = sqlFun.GetData("SELECT soNamDat,maKhenThuongpr_sd,toanTu,soNamXet FROM tblCauHinhTDKTCT WHERE sttCauHinhpr_sd = N'" + tblCauHinh.Rows[0]["sttCauHinhpr"].ToString() + "'");
                            if (tblCauHinhCT_SKKN.Rows.Count > 0)//nếu có trong bảng cấu hình chi tiết thì duyệt kiểm
                            {
                                for (int a = 0; a < tblCauHinhCT_SKKN.Rows.Count; a++)
                                {
                                    if (tblCauHinhCT_SKKN.Rows[a]["toanTu"].ToString() == "Hoặc")
                                    {
                                        //skkncs
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capCS='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);
                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";
                                                return true;
                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                            }


                                        }
                                        //skkntinh
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString()) > 0)
                                        {
                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTinh='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);
                                            if (tongsoskknnamtruoc >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";
                                                return true;
                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                            }


                                        }
                                        //skkntq
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTQ='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);


                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";
                                                return true;
                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                            }


                                        }

                                    }
                                    else
                                    {
                                        //skkncs
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capCS='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);
                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";

                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                                break;
                                            }


                                        }
                                        //skkntinh
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTinh='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";

                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                                break;
                                            }


                                        }
                                        //skkntq
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTQ='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);


                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";

                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                                break;
                                            }


                                        }
                                    }
                                }
                                if (tieuchuan == "dutieuchuan")
                                    return true;
                                else
                                    tieuchuan = "khongdutieuchuan";
                            }
                            else // nếu không có trong cấu hình chi tiết mà số lượng là 1 thì kiểm đăng ký skkn năm xét
                            {
                                if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString()) == 1)
                                {
                                    SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy)) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capCS='1'";
                                    decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                    if (tongsoskknnamtruoc > 0)
                                    {
                                        tieuchuan = "dutieuchuan";//đã tìm được 
                                        return true;
                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                    }

                                }
                                if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString()) == 1)
                                {
                                    SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy)) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTinh='1'";
                                    decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                    if (tongsoskknnamtruoc > 0)
                                    {
                                        tieuchuan = "dutieuchuan";//đã tìm được 
                                        return true;
                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                    }
                                }
                                if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString()) == 1)
                                {

                                    SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy)) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTQ='1'";
                                    decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                    if (tongsoskknnamtruoc > 0)
                                    {
                                        tieuchuan = "dutieuchuan";//đã tìm được 
                                        return true;
                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                    }

                                }
                                tieuchuan = "khongdutieuchuan"; // nếu không có  đăng ký skkn năm xét thì là không đủ đk
                            }
                        }
                        //kiểm tra cấu hình chi tiết
                        DataTable tblCauHinhCT = new DataTable();
                        tblCauHinhCT = sqlFun.GetData("SELECT soNamDat,maKhenThuongpr_sd,toanTu,soNamXet FROM tblCauHinhTDKTCT WHERE sttCauHinhpr_sd = N'" + tblCauHinh.Rows[0]["sttCauHinhpr"].ToString() + "'");
                        if (tblCauHinhCT.Rows.Count > 0)//nếu có trong bảng cấu hình
                        {
                            for (int l = 0; l < tblCauHinhCT.Rows.Count; l++)
                            {
                                //1. kiểm nếu toán tử nếu hoặc thì chỉ cần một tiêu chí đủ đk thì đủ điều kiện
                                if (tblCauHinhCT.Rows[l]["toanTu"].ToString() == "Hoặc")
                                {


                                    decimal sonam = 0;
                                    //kiểm tra tồn tại danh hiệu đang xét của các năm trước nếu đã được công nhận thì mốc tính sẽ là mốc lúc công nhận
                                    string sqllaynamdatdanhhieu = @"SELECT  convert(decimal,YEAR(ngayKyQD)) FROM tblThiDuaKhenThuong WHERE  (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "') AND maKhenThuongpr_sd ='" + maKhenThuong + "'";
                                    decimal laynamdatdanhhieu = sqlFun.GetOneDecimalField(sqllaynamdatdanhhieu);
                                    string sqltongsonamtruoc = "";
                                    if (laynamdatdanhhieu != 0)
                                    {
                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + laynamdatdanhhieu + "' AND '" + namXetDangKy + "')";

                                    }
                                    else
                                    {

                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "')";

                                    }
                                    decimal tongsonamtruoc = sqlFun.GetOneDecimalField(sqltongsonamtruoc);
                                    //lấy đăng ký của năm hiện tại
                                    SqlFunction sqlFun_ = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqlsonam_hientai = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + @"' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + @"' AND YEAR(ngayDangKy) ='" + namXetDangKy + @"' and (ngayKyQD is  null or ngayKyQD ='') ";
                                    decimal sonam_hientai = sqlFun_.GetOneDecimalField(sqlsonam_hientai);
                                    if (dangKyExcel != "")
                                    {
                                        sonam_hientai = 1;
                                    }
                                    sonam = tongsonamtruoc;
                                    if (sonam >= (Convert.ToDecimal(tblCauHinhCT.Rows[l]["soNamDat"].ToString())))
                                    {
                                        tieuchuan = "dutieuchuan";
                                        return true;
                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                    }


                                }
                                else //ngược lại là toán tử và thì bắt buộc bảng chi tiết phải đảm bảo đủ điều kiện
                                {

                                    decimal sonam = 0;
                                    //kiểm tra tồn tại danh hiệu đang xét của các năm trước nếu đã được công nhận thì mốc tính sẽ là mốc lúc công nhận
                                    string sqllaynamdatdanhhieu = @"SELECT  convert(decimal,YEAR(ngayKyQD)) FROM tblThiDuaKhenThuong WHERE  (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "') AND maKhenThuongpr_sd ='" + maKhenThuong + "'";
                                    decimal laynamdatdanhhieu = sqlFun.GetOneDecimalField(sqllaynamdatdanhhieu);
                                    string sqltongsonamtruoc = "";
                                    if (laynamdatdanhhieu != 0)
                                    {
                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + laynamdatdanhhieu + "' AND '" + namXetDangKy + "')";

                                    }
                                    else
                                    {

                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "')";

                                    }
                                    decimal tongsonamtruoc = sqlFun.GetOneDecimalField(sqltongsonamtruoc);
                                    //lấy đăng ký của năm hiện tại
                                    SqlFunction sqlFun_ = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqlsonam_hientai = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + @"' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + @"' AND YEAR(ngayDangKy) ='" + namXetDangKy + @"' and (ngayKyQD is  null or ngayKyQD ='') ";
                                    decimal sonam_hientai = sqlFun_.GetOneDecimalField(sqlsonam_hientai);
                                    if (dangKyExcel != "")
                                    {
                                        sonam_hientai = 1;
                                    }
                                    sonam = tongsonamtruoc;
                                    if (sonam >= (Convert.ToDecimal(tblCauHinhCT.Rows[l]["soNamDat"].ToString())))
                                    {
                                        tieuchuan = "dutieuchuan";

                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                        break;
                                    }


                                }
                            }
                            if (tieuchuan == "dutieuchuan")
                                return true;
                            else
                                tieuchuan = "khongdutieuchuan";
                        }
                        if (tieuchuan == "dutieuchuan")
                            return true;
                        else
                            return false;
                    }
                    else //ngược lại là toán tử và thì bắt buộc bảng chính và bảng chi tiết phải đảm bảo đủ điều kiện
                    {
                        //nếu check có skkn được check và số lương skkncs hoặc skkntinh > 1 thì xét điều kiện
                        if (tblCauHinh.Rows[0]["coSKKN"].ToString() == "True")
                        {
                            //lấy số năm xét để xét skkn
                            DataTable tblCauHinhCT_SKKN = new DataTable();
                            tblCauHinhCT_SKKN = sqlFun.GetData("SELECT soNamDat,maKhenThuongpr_sd,toanTu,soNamXet FROM tblCauHinhTDKTCT WHERE sttCauHinhpr_sd = N'" + tblCauHinh.Rows[0]["sttCauHinhpr"].ToString() + "'");
                            if (tblCauHinhCT_SKKN.Rows.Count > 0)//nếu có trong bảng cấu hình chi tiết thì duyệt kiểm
                            {
                                for (int a = 0; a < tblCauHinhCT_SKKN.Rows.Count; a++)
                                {
                                    if (tblCauHinhCT_SKKN.Rows[a]["toanTu"].ToString() == "Hoặc")
                                    {
                                        //skkncs
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capCS='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);
                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";
                                                break;
                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                            }


                                        }
                                        //skkntinh
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTinh='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);
                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";
                                                break;
                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                            }


                                        }
                                        //skkntq
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTQ='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";
                                                break;
                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                            }


                                        }

                                    }
                                    else
                                    {
                                        //skkncs
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capCS='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);


                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";

                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                                return false;
                                            }


                                        }
                                        //skkntinh
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTinh='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);


                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";

                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                                return false;
                                            }


                                        }
                                        //skkntq
                                        if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString()) > 0)
                                        {

                                            SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                            string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy) - (Convert.ToDecimal(tblCauHinhCT_SKKN.Rows[a]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTQ='1'";
                                            decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                            if ((tongsoskknnamtruoc) >= (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString())))
                                            {
                                                tieuchuan = "dutieuchuan";

                                            }
                                            else
                                            {
                                                tieuchuan = "khongdutieuchuan";
                                                return false;
                                            }


                                        }
                                    }
                                }
                                if (tieuchuan == "dutieuchuan")
                                    tieuchuan = "dutieuchuan";
                                else
                                    return false;
                            }
                            else // nếu không có trong cấu hình chi tiết mà số lượng là 1 thì kiểm đăng ký skkn năm xét
                            {
                                if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNCS"].ToString()) == 1)
                                {
                                    SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy)) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capCS='1'";
                                    decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                    if (tongsoskknnamtruoc > 0)
                                    {
                                        tieuchuan = "dutieuchuan";//đã tìm được 

                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                        return false;
                                    }

                                }
                                if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTinh"].ToString()) == 1)
                                {
                                    SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy)) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTinh='1'";
                                    decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                    if (tongsoskknnamtruoc > 0)
                                    {
                                        tieuchuan = "dutieuchuan";//đã tìm được 

                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                        return false;
                                    }

                                }
                                if (Convert.ToDecimal(tblCauHinh.Rows[0]["SKKNTQ"].ToString()) == 1)
                                {
                                    SqlFunction sqlFunSKKNCS = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqltongsoskknnamtruoc = @"select convert(decimal(18,0),count(sttSKKNpr)) as sttSKKNpr from tblSangKienKN where ketQua=N'Đ' and sttDoiTuongTDKTpr_sd='" + sttDoiTuongTDKT + "' and (YEAR(ngayQuyetDinh) between '" + (Convert.ToDecimal(namXetDangKy)) + "' and '" + Convert.ToDecimal(namXetDangKy) + "') and capTQ='1'";
                                    decimal tongsoskknnamtruoc = sqlFunSKKNCS.GetOneDecimalField(sqltongsoskknnamtruoc);

                                    if (tongsoskknnamtruoc > 0)
                                    {
                                        tieuchuan = "dutieuchuan";//đã tìm được 

                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                        return false;
                                    }
                                }
                                tieuchuan = "khongdutieuchuan"; // nếu không có  đăng ký skkn năm xét thì là không đủ đk
                                return false;
                            }
                        }
                        //kiểm tra cấu hình chi tiết
                        DataTable tblCauHinhCT = new DataTable();
                        tblCauHinhCT = sqlFun.GetData("SELECT soNamDat,maKhenThuongpr_sd,toanTu,soNamXet FROM tblCauHinhTDKTCT WHERE sttCauHinhpr_sd = N'" + tblCauHinh.Rows[0]["sttCauHinhpr"].ToString() + "'");
                        if (tblCauHinhCT.Rows.Count > 0)//nếu có trong bảng cấu hình
                        {
                            for (int l = 0; l < tblCauHinhCT.Rows.Count; l++)
                            {
                                //1. kiểm nếu toán tử nếu hoặc thì chỉ cần một tiêu chí đủ đk thì đủ điều kiện
                                if (tblCauHinhCT.Rows[l]["toanTu"].ToString() == "Hoặc")
                                {


                                    decimal sonam = 0;
                                    //kiểm tra tồn tại danh hiệu đang xét của các năm trước nếu đã được công nhận thì mốc tính sẽ là mốc lúc công nhận
                                    string sqllaynamdatdanhhieu = @"SELECT  convert(decimal,YEAR(ngayKyQD)) FROM tblThiDuaKhenThuong WHERE  (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "') AND maKhenThuongpr_sd ='" + maKhenThuong + "'";
                                    decimal laynamdatdanhhieu = sqlFun.GetOneDecimalField(sqllaynamdatdanhhieu);
                                    string sqltongsonamtruoc = "";
                                    if (laynamdatdanhhieu != 0)
                                    {
                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + laynamdatdanhhieu + "' AND '" + namXetDangKy + "')";

                                    }
                                    else
                                    {

                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "')";

                                    }
                                    decimal tongsonamtruoc = sqlFun.GetOneDecimalField(sqltongsonamtruoc);
                                    //lấy đăng ký của năm hiện tại
                                    SqlFunction sqlFun_ = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqlsonam_hientai = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + @"' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + @"' AND YEAR(ngayDangKy) ='" + namXetDangKy + @"' and (ngayKyQD is  null or ngayKyQD ='') ";
                                    decimal sonam_hientai = sqlFun_.GetOneDecimalField(sqlsonam_hientai);
                                    if (dangKyExcel != "")
                                    {
                                        sonam_hientai = 1;
                                    }
                                    sonam = tongsonamtruoc;
                                    if (sonam >= (Convert.ToDecimal(tblCauHinhCT.Rows[l]["soNamDat"].ToString())))
                                    {
                                        tieuchuan = "dutieuchuan";
                                        break;
                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                    }


                                }
                                else //ngược lại là toán tử và thì bắt buộc bảng chi tiết phải đảm bảo đủ điều kiện
                                {

                                    decimal sonam = 0;
                                    //kiểm tra tồn tại danh hiệu đang xét của các năm trước nếu đã được công nhận thì mốc tính sẽ là mốc lúc công nhận
                                    string sqllaynamdatdanhhieu = @"SELECT  convert(decimal,YEAR(ngayKyQD)) FROM tblThiDuaKhenThuong WHERE  (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "') AND maKhenThuongpr_sd ='" + maKhenThuong + "'";
                                    decimal laynamdatdanhhieu = sqlFun.GetOneDecimalField(sqllaynamdatdanhhieu);
                                    string sqltongsonamtruoc = "";
                                    if (laynamdatdanhhieu != 0)
                                    {
                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + laynamdatdanhhieu + "' AND '" + namXetDangKy + "')";

                                    }
                                    else
                                    {

                                        sqltongsonamtruoc = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + "' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + "' AND (YEAR(ngayKyQD) BETWEEN '" + (Convert.ToInt32(namXetDangKy) - (Convert.ToInt32(tblCauHinhCT.Rows[l]["soNamXet"].ToString()) - (namhanhchinh == 1 ? 1 : 1))) + "' AND '" + namXetDangKy + "')";

                                    }
                                    decimal tongsonamtruoc = sqlFun.GetOneDecimalField(sqltongsonamtruoc);
                                    //lấy đăng ký của năm hiện tại
                                    SqlFunction sqlFun_ = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
                                    string sqlsonam_hientai = @"SELECT  convert(decimal,COUNT(sttTDKTpr)) FROM tblThiDuaKhenThuong 
                                        WHERE sttDoiTuongTDKTpr_sd = N'" + sttDoiTuongTDKT + @"' and maKhenThuongpr_sd = '" + tblCauHinhCT.Rows[l]["maKhenThuongpr_sd"].ToString() + @"' AND YEAR(ngayDangKy) ='" + namXetDangKy + @"' and (ngayKyQD is  null or ngayKyQD ='') ";
                                    decimal sonam_hientai = sqlFun_.GetOneDecimalField(sqlsonam_hientai);
                                    if (dangKyExcel != "")
                                    {
                                        sonam_hientai = 1;
                                    }
                                    sonam = tongsonamtruoc;
                                    if (sonam >= (Convert.ToDecimal(tblCauHinhCT.Rows[l]["soNamDat"].ToString())))
                                    {
                                        tieuchuan = "dutieuchuan";

                                    }
                                    else
                                    {
                                        tieuchuan = "khongdutieuchuan";
                                        return false;
                                    }


                                }
                            }
                            if (tieuchuan == "khongdutieuchuan")
                                return false;
                            else
                                tieuchuan = "dutieuchuan";
                        }
                        if (tieuchuan == "dutieuchuan")
                            return true;
                        else
                            return false;
                    }

                }
                else
                {
                    tieuchuan = "khongdutieuchuan";
                    return false;
                }

                //if (tieuchuan == "khongdutieuchuan")
                //{

                //    return false;
                //}
                //else
                //{

                //    return true;
                //}
            }
            catch
            {
                return false;
            }
        }
        [AjaxPro.AjaxMethod]
        public bool LuuCaNhan(object[] para)
        {
            SqlConnection sqlConn = new SqlConnection(HttpContext.Current.Session.GetConnectionString2());
            try
            {
                sqlConn.Open();
                SqlCommand sqlCom = new SqlCommand("INSERT INTO dbo.tblDMDoiTuongTDKT( tenDoiTuongTDKT ,tenChucVuQL ,gioiTinh ,ngaySinh,namSinh ,CMND ,noiCap ,ngayCap ,loaiDoiTuong ,sttDoiTuongTDKTpr_cha,danhXungpr_sd,maDoiTuong ,maDonVipr_sd ,ngayThaoTac ,nguoiThaoTac,laLanhDao,ngheThuatXiecMua,maDonVipr_noicongtac,ngayVeCQ, ngungSD)"
                     + " VALUES(@tenDoiTuongTDKT ,@tenChucVuQL ,@gioiTinh ,@ngaySinh,@namSinh ,@CMND ,@noiCap ,@ngayCap ,@loaiDoiTuong ,@sttDoiTuongTDKTpr_cha,@danhXungpr_sd,@maDoiTuong,@maDonVipr_sd ,@ngayThaoTac ,@nguoiThaoTac,@laLanhDao,@ngheThuatXiecMua,@maDonVipr_noicongtac,@ngayVeCQ,@ngungSD)", sqlConn);
                sqlCom.Parameters.Add(new SqlParameter("@tenDoiTuongTDKT", para[0]));
                sqlCom.Parameters.Add(new SqlParameter("@maDonVipr_noicongtac", para[11]));
                sqlCom.Parameters.Add(new SqlParameter("@tenChucVuQL", para[1].ToString()));
                if (string.IsNullOrEmpty(para[12].ToString()))
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngayVeCQ", DBNull.Value));
                }
                else
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngayVeCQ", DateTime.Parse(para[12].ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb"))));
                }
                sqlCom.Parameters.Add(new SqlParameter("@gioiTinh", para[2].ToString()));
                //sqlCom.Parameters.Add(new SqlParameter("@ngaySinh", DateTime.Parse(para[3].ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb"))));
                sqlCom.Parameters.Add(new SqlParameter("@CMND", para[4].ToString()));
                if (string.IsNullOrEmpty(para[3].ToString()))
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngaySinh", DBNull.Value));
                    sqlCom.Parameters.Add(new SqlParameter("@namSinh", "false"));

                }
                else
                {
                    if (para[3].ToString().Length < 5)
                    {
                        sqlCom.Parameters.Add(new SqlParameter("@namSinh", "true"));
                        sqlCom.Parameters.Add(new SqlParameter("@ngaySinh", DateTime.Parse("01/01/" + para[3].ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb"))));
                    }
                    else
                    {
                        sqlCom.Parameters.Add(new SqlParameter("@namSinh", "false"));
                        sqlCom.Parameters.Add(new SqlParameter("@ngaySinh", DateTime.Parse(para[3].ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb"))));
                    }
                }
                if (string.IsNullOrEmpty(para[5].ToString()))
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngayCap", DBNull.Value));
                }
                else
                {
                    sqlCom.Parameters.Add(new SqlParameter("@ngayCap", DateTime.Parse(para[5].ToString(), System.Globalization.CultureInfo.GetCultureInfo("en-gb"))));
                }
                sqlCom.Parameters.Add(new SqlParameter("@noiCap", para[6].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@loaiDoiTuong", "Cá nhân"));

                if (string.IsNullOrEmpty(para[7].ToString()))
                {
                    sqlCom.Parameters.Add(new SqlParameter("@sttDoiTuongTDKTpr_cha", DBNull.Value));
                }
                else
                {
                    sqlCom.Parameters.Add(new SqlParameter("@sttDoiTuongTDKTpr_cha", para[7].ToString()));
                }
                sqlCom.Parameters.Add(new SqlParameter("@danhXungpr_sd", para[8].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@maDoiTuong", para[9].ToString()));

                sqlCom.Parameters.Add(new SqlParameter("@maDonVipr_sd", HttpContext.Current.Session.GetMaDonVi()));
                sqlCom.Parameters.Add(new SqlParameter("@ngayThaoTac", HttpContext.Current.Session.GetCurrentDatetimeMMddyyyy("MM/dd/yyyy")));
                sqlCom.Parameters.Add(new SqlParameter("@nguoiThaoTac", HttpContext.Current.Session.GetCurrentUserID()));
                sqlCom.Parameters.Add(new SqlParameter("@laLanhDao", para[10].ToString()));
                sqlCom.Parameters.Add(new SqlParameter("@ngheThuatXiecMua", "0"));
                sqlCom.Parameters.Add(new SqlParameter("@ngungSD", "0"));
                if (sqlCom.ExecuteNonQuery() > 0)
                {
                    sqlConn.Close(); return true;
                }
                else
                {
                    sqlConn.Close(); return false;
                }

            }
            catch
            {
                sqlConn.Close(); return false;
            }
        }

        [AjaxPro.AjaxMethod]
        public string XuatWord(object sttQuyetDinhKTpr)
        {
            SqlFunction _SqlClass = new SqlFunction(HttpContext.Current.Session.GetConnectionString2());
            string url = "~/xuatword" + "/" + HttpContext.Current.Session.GetCurrentUserID() + "/";
            string fileName = "BK-26TT-54CN-10nLuatPCTN_R" + HttpContext.Current.Session.GetMaDonVi() + "-" + (DateTime.Now.ToString("ddMMyyyyHHmmss")) + ".docx";

            if (!System.IO.Directory.Exists(Server.MapPath(url)))
            {
                System.IO.Directory.CreateDirectory(Server.MapPath(url));
            }
            DirectoryInfo di = new DirectoryInfo(Server.MapPath(url));
            FileInfo[] rgFiles = di.GetFiles();
            foreach (FileInfo fi in rgFiles)
            {
                fi.Delete();
            }

            File.Copy(Server.MapPath("~/WordMau/BK-26TT-54CN-10nLuatPCTN_R.docx"), Server.MapPath("~/xuatword" + "/" + HttpContext.Current.Session.GetCurrentUserID() + "/" + fileName), true);

            #region truy van
            string sql = @"SELECT soQuyetDinh,ngayKhenThuong=CONVERT(NVARCHAR(10),ngayKhenThuong,103)
                           ,soLuong=(SELECT COUNT(sttTDKTpr) FROM dbo.tblThiDuaKhenThuong WHERE sttQuyetDinhKTpr_sd=tblQuyetDinhKT.sttQuyetDinhKTpr)
                           ,danhHieu=(SELECT tenKhenThuong FROM dbo.tblDMKhenThuong WHERE maKhenThuongpr=tblQuyetDinhKT.maKhenThuongpr_sd)
                           FROM dbo.tblQuyetDinhKT WHERE sttQuyetDinhKTpr=N'" + sttQuyetDinhKTpr.ToString() + "'";
            #endregion
            var nfi = new NumberFormatInfo { NumberDecimalSeparator = ",", NumberGroupSeparator = "." };
            DataTable dt = new DataTable();
            dt = _SqlClass.GetData(sql);
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(HttpContext.Current.Server.MapPath(url + fileName), true))
            {
                try
                {
                    MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                    Body body = wordDoc.MainDocumentPart.Document.Body;
                    foreach (var text in body.Descendants<Text>())
                    {
                        if (dt.Rows.Count > 0)
                        {
                            string ngay = dt.Rows[0]["ngayKhenThuong"].ToString();
                            if (text.Text.Contains("soQuyetDinh"))
                            {
                                text.Text = text.Text.Replace("soQuyetDinh", dt.Rows[0]["soQuyetDinh"].ToString());
                            }
                            if (text.Text.Contains("pDiaDanh"))
                            {
                                text.Text = text.Text.Replace("pDiaDanh", HttpContext.Current.Session.GetDiadanh() + ", ngày " + ngay.Substring(0, 2) + " tháng " + ngay.Substring(3, 2) + " năm " + ngay.Substring(6, 4));
                            }
                            if (text.Text.Contains("pDonVi"))
                            {
                                text.Text = text.Text.Replace("pDonVi", HttpContext.Current.Session.GetTenDonVi());
                            }
                            if (text.Text.Contains("kyTen"))
                            {
                                text.Text = text.Text.Replace("kyTen", "");
                            }
                            if (text.Text.Contains("pCapTren"))
                            {
                                text.Text = text.Text.Replace("pCapTren", HttpContext.Current.Session.GetDonViCapTren());
                            }
                        }
                    }
                    var paras = body.Elements();
                    foreach (var para1 in paras)
                    {
                        foreach (var run1 in para1.Elements<Run>())
                        {
                            foreach (var text in run1.Elements<Text>())
                            {
                                if (dt.Rows.Count > 0)
                                {
                                    if (text.Text.Contains("fDoiTuong"))
                                    {
                                        text.Text = text.Text.Replace("fDoiTuong", dt.Rows[0]["soLuong"].ToString() + " cá nhân");
                                    }
                                    if (text.Text.Contains("_soLuong"))
                                    {
                                        text.Text = text.Text.Replace("_soLuong", dt.Rows[0]["soLuong"].ToString() + " cá nhân");
                                    }
                                    if (text.Text.Contains("fSoLuong"))
                                    {
                                        text.Text = text.Text.Replace("fSoLuong", dt.Rows[0]["soLuong"].ToString() + " cá nhân").ToUpper();
                                    }
                                    if (text.Text.Contains("danhHieu"))
                                    {
                                        text.Text = text.Text.Replace("danhHieu", dt.Rows[0]["danhHieu"].ToString());
                                    }
                                    if (text.Text.Contains("fDanhHieu"))
                                    {
                                        text.Text = text.Text.Replace("fDanhHieu", dt.Rows[0]["danhHieu"].ToString());
                                    }
                                }

                            }
                        }
                    }

                    var tables = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Table>().ToList();
                    DataTable dt1 = _SqlClass.GetData(@"SELECT hoTen=(SELECT tenDoiTuongTDKT FROM dbo.tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr=tblThiDuaKhenThuong.sttDoiTuongTDKTpr_sd)
                                              ,chucVu=(SELECT tenChucVuQL FROM dbo.tblDMDoiTuongTDKT WHERE sttDoiTuongTDKTpr=tblThiDuaKhenThuong.sttDoiTuongTDKTpr_sd)
                                              ,donVi=(SELECT tenDonVi FROM dbo.tblDMDonvi WHERE maDonvipr=tblThiDuaKhenThuong.maDonVipr_noicongtac)
                                              FROM dbo.tblThiDuaKhenThuong WHERE sttQuyetDinhKTpr_sd IN (SELECT sttQuyetDinhKTpr FROM dbo.tblQuyetDinhKT WHERE sttQuyetDinhKTpr=N'" + sttQuyetDinhKTpr.ToString() + "')");
                    if (dt1.Rows.Count > 0)
                    {
                        int stt = 0;
                        var trTD = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                        var td1TD = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                        var td2TD = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                        td1TD.Append(new TableCellProperties(
                               new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }),
                               new Paragraph(
                                      new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                                     new ParagraphProperties(new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact, Before = "120", After = "120" }),
                                     new DocumentFormat.OpenXml.Wordprocessing.Run(
                                     new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                     new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }),
                                     new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Bold { Val = true }),
                                     new DocumentFormat.OpenXml.Wordprocessing.Text("I.")
                                     )
                                 )
                             );

                        td2TD.Append(new TableCellProperties(
                            new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }),
                            new Paragraph(
                                new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                                new ParagraphProperties(new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact, Before = "120", After = "120" }),
                                new DocumentFormat.OpenXml.Wordprocessing.Run(
                                    new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                    new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }),
                                    new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Bold { Val = true }),
                                    new DocumentFormat.OpenXml.Wordprocessing.Text("Cá nhân")
                                    )
                                )
                            );
                        trTD.Append(td1TD, td2TD);
                        tables[2].Append(trTD);
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            stt += 1;
                            var tr = new DocumentFormat.OpenXml.Wordprocessing.TableRow();
                            var td1 = new DocumentFormat.OpenXml.Wordprocessing.TableCell();
                            var td2 = new DocumentFormat.OpenXml.Wordprocessing.TableCell();

                            td1.Append(new TableCellProperties(
                                     new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }),
                                     new Paragraph(
                                         new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                                         new ParagraphProperties(new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact, Before = "120", After = "120" }),
                                         new DocumentFormat.OpenXml.Wordprocessing.Run(
                                             new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                             new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }),
                                             new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Bold { Val = false }),
                                             new DocumentFormat.OpenXml.Wordprocessing.Text(stt.ToString() + ".")
                                            )
                                         )
                                     );

                            td2.Append(new TableCellProperties(
                                new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }),
                                new Paragraph(
                                    new ParagraphProperties(new Justification() { Val = JustificationValues.Left }),
                                    new ParagraphProperties(new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact, Before = "120", After = "120" }),
                                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                                        new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                                        new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "28" }),
                                        new RunProperties(new DocumentFormat.OpenXml.Wordprocessing.Italic { Val = false }),
                                        new DocumentFormat.OpenXml.Wordprocessing.Text(dt1.Rows[i]["hoTen"].ToString() + ", " + (dt1.Rows[i]["chucVu"].ToString() == "" ? "" : dt1.Rows[i]["chucVu"].ToString() + " ") + dt1.Rows[i]["donVi"].ToString())
                                        )
                                    )
                                );
                            tr.Append(td1, td2);
                            tables[2].Append(tr);
                        }
                    }

                    wordDoc.MainDocumentPart.Document.Save();
                    wordDoc.Close();
                    return "/xuatword/" + HttpContext.Current.Session.GetCurrentUserID() + "/" + fileName;
                }
                catch (Exception ex)
                {
                    wordDoc.MainDocumentPart.Document.Save();
                    wordDoc.Close();
                    return "/xuatword/" + HttpContext.Current.Session.GetCurrentUserID() + "/" + fileName;
                }
            }

        }
    }
}
