<!--#include file="conn.asp"-->
//运行文本域代码
function runCode(obj) {
        var winname = window.open('', "_blank", '');
        winname.document.open('text/html', 'replace');
	winname.opener = null // 防止代码对页面修改
        winname.document.write(obj.value);
        winname.document.close();
}
function saveCode(obj) {
        var winname = window.open('', '_blank', 'top=10000');
        winname.document.open('text/html', 'replace');
        winname.document.write(obj.value);
        winname.document.execCommand('saveas','','code.htm');
        winname.close();
}
function copycode(obj) {
	obj.select(); 
	js=obj.createTextRange(); 
	js.execCommand("Copy")
}

//图片缩放
function resizeimg(ImgD,iwidth,iheight) {
     var image=new Image();
     image.src=ImgD.src;
     if(image.width>0 && image.height>0){
        if(image.width/image.height>= iwidth/iheight){
           if(image.width>iwidth){
               ImgD.width=iwidth;
               ImgD.height=(image.height*iwidth)/image.width;
           }else{
                  ImgD.width=image.width;
                  ImgD.height=image.height;
                }
               ImgD.alt=image.width+"×"+image.height;
        }
        else{
                if(image.height>iheight){
                       ImgD.height=iheight;
                       ImgD.width=(image.width*iheight)/image.height;
                }else{
                        ImgD.width=image.width;
                        ImgD.height=image.height;
                     }
                ImgD.alt=image.width+"×"+image.height;
            }
　　　　　ImgD.style.cursor= "pointer"; //改变鼠标指针
　　　　　ImgD.onclick = function() { window.open(this.src);} //点击打开大图片
　　　　if (navigator.userAgent.toLowerCase().indexOf("ie") > -1) { //判断浏览器，如果是IE
　　　　　　ImgD.title = "请使用鼠标滚轮缩放图片!";
　　　　　　ImgD.onmousewheel = function img_zoom() //滚轮缩放
　　　　　 {
　　　　　　　　　　var zoom = parseInt(this.style.zoom, 10) || 100;
　　　　　　　　　　zoom += event.wheelDelta / 12;
　　　　　　　　　　if (zoom> 0)　this.style.zoom = zoom + "%";
　　　　　　　　　　return false;
　　　　　 }
　　　  } else { //如果不是IE
　　　　　　　     ImgD.title = "点击图片可在新窗口打开";
　　　　　　   }
    }
}

//双击鼠标滚动屏幕的代码
var currentpos,timer; 
function initialize() 
{ 
timer=setInterval("scrollwindow()",16); 
} 
function sc(){ 
clearInterval(timer); 
} 
function scrollwindow() 
{ 
currentpos=document.documentElement.scrollTop; 
window.scroll(0,++currentpos); 
if (currentpos != document.documentElement.scrollTop) 
sc(); 
} 
document.onmousedown=sc 
document.ondblclick=initialize


function Getcolor(img_val,input_val){
	var arr = showModalDialog("../xwskin/selcolor.html?action=title", "", "dialogWidth:18.5em; dialogHeight:17.5em; status:0; help:0");
	if (arr != null){
		document.getElementById(input_val).value = arr;
		img_val.style.backgroundColor = arr;
		}
}

function SetCookie(name,value){
    var argv=SetCookie.arguments;
    var argc=SetCookie.arguments.length;
    var expires=(2<argc)?argv[2]:null;
    var path=(3<argc)?argv[3]:null;
    var domain=(4<argc)?argv[4]:null;
    var secure=(5<argc)?argv[5]:false;
    document.cookie=name+"="+escape(value)+((expires==null)?"":("; expires="+expires.toGMTString()))+((path==null)?"":("; path="+path))+((domain==null)?"":("; domain="+domain))+((secure==true)?"; secure":""); 
}

function GetCookie(Name) {
    var search = Name + "=";
    var returnvalue = "";
    if (document.cookie.length > 0) {
          offset = document.cookie.indexOf(search);
          if (offset != -1) {      
                offset += search.length;
                end = document.cookie.indexOf(";", offset);                        
                if (end == -1)
                      end = document.cookie.length;
                returnvalue=unescape(document.cookie.substring(offset,end));
          }
    }
    return returnvalue;
}
    
function changecss(url){
    if(url!=""){
          skin.href=url;
          var expdate=new Date();
          expdate.setTime(expdate.getTime()+(24*60*60*1000*30));
          //expdate=null;
                                  //以下设置COOKIES时间为1年,自己随便设置该时间..
          SetCookie("nowskin",url,expdate,"/",null,false);
    }
}

var flag=false;
function DrawImage(ImgD){
var image=new Image();
var iwidth = 600;
//这里设置最大高度 var iheight = 450; 
image.src=ImgD.src;
if(image.width>0 && image.height>0){
   flag=true;
   if(image.width/image.height>= iwidth/iheight){
    if(image.width>iwidth){ 
     ImgD.width=iwidth;
     ImgD.height=(image.height*iwidth)/image.width;
    }else{
     ImgD.width=image.width; 
     ImgD.height=image.height;
    }
   }else{
    if(image.height>iheight){ 
     ImgD.height=iheight;
     ImgD.width=(image.width*iheight)/image.height; 
    }else{
     ImgD.width=image.width; 
     ImgD.height=image.height;
    }
   }
}
}

function addfavorite()
{
 if (document.all)
 {
 window.external.addFavorite('http://<%=SiteUrl%>','<%=SiteTitle%>');
 }
 else if (window.sidebar)
 {
 window.sidebar.addPanel('<%=SiteTitle%>', 'http://<%=SiteUrl%>', "");
 }
} 

function getvote(id){
    GETAjax("inc/show_vote.asp?id="+id,"votebar_"+id);
    return false;
}

var qi;var qt;var qp="parentNode";var qc="className";function ldc(sd,v,l){if(!l){l=1;sd=document.getElementById("ld"+sd);sd.onmouseover=function(e){x6(e)};document.onmouseover=x2;sd.style.zoom=1;}sd.style.zIndex=l;var lsp;var sp=sd.childNodes;for(var i=0;i<sp.length;i++){var b=sp[i];if(b.tagName=="A"){lsp=b;b.onmouseover=x0;if(l==1&&v){b.style.styleFloat="none";b.style.cssFloat="none";}}if(b.tagName=="DIV"){if(window.showHelp&&!window.XMLHttpRequest)sp[i].insertAdjacentHTML("afterBegin","<span style='display:block;font-size:1px;height:0px;width:0px;visibility:hidden;'></span>");x5("ldparent",lsp,1);lsp.cdiv=b;b.idiv=lsp;new ldc(b,null,l+1);}}};function x2(e){if(qi&&!qt)qt=setTimeout("x3()",100);};function x3(){var a;if((a=qi)){do{x1(a);}while((a=a[qp])&&!ld_a(a))}qi=null;};function ld_a(a){if(a[qc].indexOf("ldmc")+1)return 1;};function x1(a){if(window.ldad&&ldad.bhide)eval(ldad.bhide);a.style.visibility="";x5("ldactive",a.idiv);};function x0(e){if(qt){clearTimeout(qt);qt=null;}var a=this;if(a[qp].isrun)return;var go=true;while((a=a[qp])&&!ld_a(a)){if(a==qi)go=false;}if(qi&&go){a=this;if((!a.cdiv)||(a.cdiv&&a.cdiv!=qi))x1(qi);a=qi;while((a=a[qp])&&!ld_a(a)){if(a!=this[qp])x1(a);else break;}}var b=this;if(b.cdiv){var aw=b.offsetWidth;var ah=b.offsetHeight;var ax=b.offsetLeft;var ay=b.offsetTop;if(ld_a(b[qp])&&b.style.styleFloat!="none"&&b.style.cssFloat!="none")aw=0;else ah=0;if(!b.cdiv.ismove){b.cdiv.style.left=(ax+aw)+"px";b.cdiv.style.top=(ay+ah)+"px";}x5("ldactive",this,1);if(window.ldad&&ldad.bvis)eval(ldad.bvis);b.cdiv.style.visibility="inherit";qi=b.cdiv;}else  if(!ld_a(b[qp]))qi=b[qp];else qi=null;x6(e);};function x5(name,b,add){var a=b[qc];if(add){if(a.indexOf(name)==-1)b[qc]+=(a?' ':'')+name;}else {b[qc]=a.replace(" "+name,"");b[qc]=b[qc].replace(name,"");}};function x6(e){if(!e)e=event;e.cancelBubble=true;if(e.stopPropagation)e.stopPropagation();}