using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using ZaloCSharpSDK;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NTSoftReceipt.zalo
{
    public partial class zalo1 : System.Web.UI.Page
    {
        public string code;
        public string json_dsbb;
        public string tranghientai;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                try
                {
                    tranghientai = "http://localhost:18140/zalo/zalo1.aspx";
                    Zalo3rdAppClient appClient = new Zalo3rdAppClient(new Zalo3rdAppInfo(2127518896466871150, "OUKlCdL8uYLIMe73QB65", tranghientai));
                    string loginUrl = appClient.getLoginUrl();
                    code = Request.QueryString["code"];
                    JObject token = appClient.getAccessToken(code);
                    if (string.IsNullOrEmpty(code))
                        Response.Redirect(loginUrl.ToString());
                    else
                    {
                        string chuoitoken = token.Property("access_token").Value.ToString();
                        JObject profile = appClient.getProfile(chuoitoken, "name,picture,id");
                        txtHDinfo.Value = profile.ToString();
                        txtTenNguoiDung.Text = profile.Property("name").Value.ToString();

                        JObject invitableFriends = appClient.getInvitableFriends(chuoitoken, 0, 0, "id, name, picture, gender");
                        int soluong = 0;
                        soluong = int.Parse(invitableFriends["summary"]["total_count"].ToString());

                        JObject banbe = appClient.getInvitableFriends(chuoitoken, 0, soluong, "id, name, picture, gender");
                        json_dsbb = banbe["data"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    Response.Write(ex.Message);
                }
            }
        }
    }
}