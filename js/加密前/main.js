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
　　　　　ImgD.style.cursor= "pointer"; 
　　　　　ImgD.onclick = function() { window.open(this.src);} 
　　　　if (navigator.userAgent.toLowerCase().indexOf("ie") > -1) { 
　　　　　　ImgD.title = "请使用鼠标滚轮缩放图片!";
　　　　　　ImgD.onmousewheel = function img_zoom() 
　　　　　 {
　　　　　　　　　　var zoom = parseInt(this.style.zoom, 10) || 100;
　　　　　　　　　　zoom += event.wheelDelta / 12;
　　　　　　　　　　if (zoom> 0)　this.style.zoom = zoom + "%";
　　　　　　　　　　return false;
　　　　　 }
　　　  } else { 
　　　　　　　     ImgD.title = "点击图片可在新窗口打开";
　　　　　　   }
    }
}