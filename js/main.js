<script>document.write(unescape('//%u8FD0%u884C%u6587%u672C%u57DF%u4EE3%u7801%0Afunction%20runCode%28obj%29%20%7B%0A%20%20%20%20%20%20%20%20var%20winname%20%3D%20window.open%28%27%27%2C%20%22_blank%22%2C%20%27%27%29%3B%0A%20%20%20%20%20%20%20%20winname.document.open%28%27text/html%27%2C%20%27replace%27%29%3B%0A%09winname.opener%20%3D%20null%20//%20%u9632%u6B62%u4EE3%u7801%u5BF9%u9875%u9762%u4FEE%u6539%0A%20%20%20%20%20%20%20%20winname.document.write%28obj.value%29%3B%0A%20%20%20%20%20%20%20%20winname.document.close%28%29%3B%0A%7D%0Afunction%20saveCode%28obj%29%20%7B%0A%20%20%20%20%20%20%20%20var%20winname%20%3D%20window.open%28%27%27%2C%20%27_blank%27%2C%20%27top%3D10000%27%29%3B%0A%20%20%20%20%20%20%20%20winname.document.open%28%27text/html%27%2C%20%27replace%27%29%3B%0A%20%20%20%20%20%20%20%20winname.document.write%28obj.value%29%3B%0A%20%20%20%20%20%20%20%20winname.document.execCommand%28%27saveas%27%2C%27%27%2C%27code.htm%27%29%3B%0A%20%20%20%20%20%20%20%20winname.close%28%29%3B%0A%7D%0Afunction%20copycode%28obj%29%20%7B%0A%09obj.select%28%29%3B%20%0A%09js%3Dobj.createTextRange%28%29%3B%20%0A%09js.execCommand%28%22Copy%22%29%0A%7D%0A%0A//%u56FE%u7247%u7F29%u653E%0Afunction%20resizeimg%28ImgD%2Ciwidth%2Ciheight%29%20%7B%0A%20%20%20%20%20var%20image%3Dnew%20Image%28%29%3B%0A%20%20%20%20%20image.src%3DImgD.src%3B%0A%20%20%20%20%20if%28image.width%3E0%20%26%26%20image.height%3E0%29%7B%0A%20%20%20%20%20%20%20%20if%28image.width/image.height%3E%3D%20iwidth/iheight%29%7B%0A%20%20%20%20%20%20%20%20%20%20%20if%28image.width%3Eiwidth%29%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.width%3Diwidth%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.height%3D%28image.height*iwidth%29/image.width%3B%0A%20%20%20%20%20%20%20%20%20%20%20%7Delse%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.width%3Dimage.width%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.height%3Dimage.height%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.alt%3Dimage.width+%22%D7%22+image.height%3B%0A%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20else%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20if%28image.height%3Eiheight%29%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.height%3Diheight%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.width%3D%28image.width*iheight%29/image.height%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%7Delse%7B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.width%3Dimage.width%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.height%3Dimage.height%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%7D%0A%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20%20ImgD.alt%3Dimage.width+%22%D7%22+image.height%3B%0A%20%20%20%20%20%20%20%20%20%20%20%20%7D%0A%u3000%u3000%u3000%u3000%u3000ImgD.style.cursor%3D%20%22pointer%22%3B%20%0A%u3000%u3000%u3000%u3000%u3000ImgD.onclick%20%3D%20function%28%29%20%7B%20window.open%28this.src%29%3B%7D%20%0A%u3000%u3000%u3000%u3000if%20%28navigator.userAgent.toLowerCase%28%29.indexOf%28%22ie%22%29%20%3E%20-1%29%20%7B%20%0A%u3000%u3000%u3000%u3000%u3000%u3000ImgD.title%20%3D%20%22%u8BF7%u4F7F%u7528%u9F20%u6807%u6EDA%u8F6E%u7F29%u653E%u56FE%u7247%21%22%3B%0A%u3000%u3000%u3000%u3000%u3000%u3000ImgD.onmousewheel%20%3D%20function%20img_zoom%28%29%20%0A%u3000%u3000%u3000%u3000%u3000%20%7B%0A%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000var%20zoom%20%3D%20parseInt%28this.style.zoom%2C%2010%29%20%7C%7C%20100%3B%0A%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000zoom%20+%3D%20event.wheelDelta%20/%2012%3B%0A%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000if%20%28zoom%3E%200%29%u3000this.style.zoom%20%3D%20zoom%20+%20%22%25%22%3B%0A%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000%u3000return%20false%3B%0A%u3000%u3000%u3000%u3000%u3000%20%7D%0A%u3000%u3000%u3000%20%20%7D%20else%20%7B%20%0A%u3000%u3000%u3000%u3000%u3000%u3000%u3000%20%20%20%20%20ImgD.title%20%3D%20%22%u70B9%u51FB%u56FE%u7247%u53EF%u5728%u65B0%u7A97%u53E3%u6253%u5F00%22%3B%0A%u3000%u3000%u3000%u3000%u3000%u3000%20%20%20%7D%0A%20%20%20%20%7D%0A%7D'));</script>