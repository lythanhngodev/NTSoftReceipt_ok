<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="zalo1.aspx.cs" Inherits="NTSoftReceipt.zalo.zalo1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>API Zalo</title>
    <link href="../public/css/css.css" rel="stylesheet" />
    <style>
        .anhdaidien{
            width:65px;
            border-radius: 50%;
            margin-right:10px;
            border: 3px solid #1aa5f6;
        }
        .hinhthe{
            width:50px;
            border-radius: 50%;
        }
        #danhsachban{
            padding: 0;
            margin: 0;
            list-style-type: none;
            height: 600px;
            overflow: auto;
        }
        #khungtinnhan{
            height:600px;
            position: relative;
        }
        #danhsachban li{
            padding: 8px;
            display: inline-table;
            width: 100%;
            float: left;
            
        }
        #danhsachban .active{
            background: #f1f1f1;
            border-left: 3px solid #1aa5f6;
        }
        #danhsachban li:hover{
            background:#f1f1f1;
            cursor: pointer;
        }
        #danhsachban img{
            width: 50px;
            border-radius: 50%;
            float: left;
        }
        #danhsachban div{
            float: left;
            width: 200px;
            font-size: 20px;
            padding: 8px;
        }
        .khungguitin{
            position:absolute;
            bottom:0;
            left:0;
            right:0;
            height:50px;
        }
        #noidungtin{
            width:100%;
            float: left;
            border-radius: 4px;
            border: 1px solid #c2c2c2;
        }
        #guitin{
            width: 80px;
            float: right;
            margin-top: 8px;
        }
    </style>
    <script>
        var jsondanhsach = <%=json_dsbb%>;
    </script>
</head>
<body style="background: #f1f1f1;">
    <form id="form1" runat="server">
        <div class="container">
            <asp:HiddenField runat="server" ID="txtHDinfo" />
            <!-- Content Wrapper. Contains page content -->
            <!-- Main content -->
            <section class="content container-fluid" style="padding-bottom: 0;">
                <section class="content-header">
                    <span style="width: 100%;">
                        <img class="anhdaidien" />
                        <asp:Label ID="txtTenNguoiDung" Font-Size="30px" runat="server"></asp:Label>
                    </span>
                </section>
            </section>
            <section class="content">
                <div class="box box-solid" id="khungthongtin" style="padding: 10px;border-radius: 15px;">
                    <div class="row" style="width:100%;">
                        <asp:Label runat="server" ID="txtLinkDangNhap"></asp:Label>
                        <h3 style="padding-left: 15px;margin-top: 10px;">DANH SÁCH BẠN BÈ</h3>
                        <hr />
                        <div class="col-md-4">
                            <div class="col-md-12" style="padding: 0; margin-bottom: 10px;">
                                <input type="text" class="form-control" placeholder="Tìm kiếm ..." value="" id="otimkiem" />
                            </div>
                            <ul id="danhsachban" class="col-md-12">
                            </ul>
                        </div>
                        <div id="khungtinnhan" class="col-md-8">
                            <h4 class="tennguoigui text-center"></h4>
                            <hr />
                            <div class="khungguitin">
                                <textarea id="noidungtin"></textarea>
                                <div class="btn btn-primary btn-sm" id="guitin">Gửi tin</div>
                            </div>
                        </div>
                    </div>
                </div>
            </section>
        </div>
    </form>
    <!-- jQuery 3 -->
    <script src="/../lte/bower_components/jquery/dist/jquery.min.js"></script>
    <!-- Bootstrap 3.3.7 -->
    <script src="/../lte/bower_components/bootstrap/dist/js/bootstrap.min.js" defer="defer"></script>
    <!-- AdminLTE App -->
    <script src="/../lte/dist/js/adminlte.min.js" defer="defer"></script>
    <script type="text/javascript" src="/../lab/js/jquery-ui.min.js" defer="defer"></script>
    <link rel="stylesheet" type="text/css" href="/../lab/css/jquery-ui.min.css">
    <script src="https://zjs.zdn.vn/zalo/sdk.js"></script>
    <script>
        Zalo.init(
            {
                version: '2.0',
                appId: '2127518896466871150',
                redirectUrl: 'http://localhost:18140/zalo/zalo1.aspx'
            }
        );
        Zalo.getLoginStatus(function(response) {
            if (response.status === 'connected') {
                Zalo.api('/me',
                  'GET',
                  {
                      fields: 'id,name'
                  },
                  function (response) { }
                );
            } else {
                Zalo.login(curentState);
            }
        });
    </script>
    <script type="text/javascript">
        function chukhongdau(alias) {
            var str = alias;
            str = str.toLowerCase();
            str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g,"a"); 
            str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g,"e"); 
            str = str.replace(/ì|í|ị|ỉ|ĩ/g,"i"); 
            str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g,"o"); 
            str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g,"u"); 
            str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g,"y"); 
            str = str.replace(/đ/g,"d");
            str = str.replace(/!|@|%|\^|\*|\(|\)|\+|\=|\<|\>|\?|\/|,|\.|\:|\;|\'|\"|\&|\#|\[|\]|~|\$|_|`|-|{|}|\||\\/g," ");
            str = str.replace(/ + /g," ");
            str = str.trim(); 
            return str;
        }
        $(document).ready(function() {
            var info = jQuery.parseJSON($("#<%=txtHDinfo.ClientID %>").val());
            $('.anhdaidien').attr("src", info.picture.data.url);
            jsondanhsach.map(function(d){
                $('#danhsachban').append('<li idu="'+d.id+'"><img class="hinhthe" src="'+d.picture.data.url+'" /><div>'+d.name+'</div></li>');
            });
        });
        $(document).on('click','li',function() {
            $('#danhsachban li').removeClass('active');
            $(this).addClass('active');
            $('.tennguoigui').text('Gửi tin nhắn đến: '+$(this).find('div').text());
            $('.tennguoigui').attr("idu",$(this).attr('idu'));
        });
        $(document).on('keyup','#otimkiem',function(){
            $('#danhsachban').find("li").remove();
            var info = jQuery.parseJSON($("#<%=txtHDinfo.ClientID %>").val());
            var key = $(this).val();
            if (jQuery.isEmptyObject(key)) {
                jsondanhsach.map(function(d){
                    $('#danhsachban').append('<li idu="'+d.id+'"><img class="hinhthe" src="'+d.picture.data.url+'" /><div>'+d.name+'</div></li>');
                });
                return false;
            }
            jsondanhsach.map(function(d){
                if (chukhongdau(d.name.toLowerCase()).search(chukhongdau(key.toLowerCase()))!=-1) {
                    $('#danhsachban').append('<li idu="'+d.id+'"><img class="hinhthe" src="'+d.picture.data.url+'" /><div>'+d.name+'</div></li>');
                }
            });
        });
        $(document).on('click','#guitin',function(){
            if ($('.tennguoigui').text().trim().length==0) {
                tbdanger('Vui lòng chọn người nhận tin');
                return false;
            }
            Zalo.api('/me/message',
                'POST',
                {
                    message:$('#noidungtin').val(),
                    link:'https://developers.zalo.me/',
                    to:$('.tennguoigui').attr("idu")
                },
                function (response) {
                    console.log(response);
                    tbinfo(response.message);
                }
              );
        });  
    </script>
<script type="text/javascript">(function(){var t;(t=jQuery).bootstrapGrowl=function(s,e){var a,o,l;switch(e=t.extend({},t.bootstrapGrowl.default_options,e),(a=t("<div>")).attr("class","bootstrap-growl alert"),e.type&&a.addClass("alert-"+e.type),e.allow_dismiss&&(a.addClass("alert-dismissible"),a.append('<button  class="close" data-dismiss="alert" type="button"><span aria-hidden="true">&#215;</span><span class="sr-only">Close</span></button>')),a.append(s),e.top_offset&&(e.offset={from:"bottom",amount:e.top_offset}),l=e.offset.amount,t(".bootstrap-growl").each(function(){return l=Math.max(l,parseInt(t(this).css(e.offset.from))+t(this).outerHeight()+e.stackup_spacing)}),(o={position:"body"===e.ele?"fixed":"absolute",margin:0,"z-index":"9999",display:"none"})[e.offset.from]=l+"px",a.css(o),"auto"!==e.width&&a.css("width",e.width+"px"),t(e.ele).append(a),e.align){case"center":a.css({left:"50%","margin-left":"-"+a.outerWidth()/2+"px"});break;case"left":a.css("left","20px");break;default:a.css("right","20px")}return a.fadeIn(),e.delay>0&&a.delay(e.delay).fadeOut(function(){return t(this).alert("close")}),a},t.bootstrapGrowl.default_options={ele:"body",type:"info",offset:{from:"bottom",amount:20},align:"right",width:250,delay:4e3,allow_dismiss:!0,stackup_spacing:10}}).call(this);</script>
<script type="text/javascript">function tbinfo(mess){$.bootstrapGrowl('<i class="fa fa-spinner fa-spin"></i>  '+mess, {type: 'info',delay: 6000});}function tbsuccess(mess){$.bootstrapGrowl('<i class="fa fa-check"></i>  '+mess, {type: 'success',delay: 2000});}function tbdanger(mess){$.bootstrapGrowl('<i class="fa fa-times"></i>  '+mess, {type: 'danger',delay: 2000});}function tban(){$('.bootstrap-growl').remove();}</script>
</body>
</html>
