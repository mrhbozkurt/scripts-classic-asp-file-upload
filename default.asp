<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
</head>

<body>
<%
process = request("process")
select case process

case "yukle"
	
	
	Server.ScriptTimeout = 1000
	
	dim upload6 
	dim count6 
	dim aciklama6 
	Set Upload6 = Server.CreateObject("Persits.Upload") 
	Upload6.OverwriteFiles = false
	Count6 = Upload6.SaveToMemory 
	dim DUzn6 'DosyaUzantisi 
	dim SunucuYeri6 
	dim YuklenecekDosya6 
	dim File6 
	
	on error resume next
	
	Set File6 = Upload6.Files(1) 
	SunucuYeri6 = Server.MapPath("dosyalar")&"\"  
	DUzn6 = UCase(Right(File6.ExtractFileName, 4))
	
	
	
	dim YuklenenDosyaninIsmi6 
	dim sql6 
	YuklenecekDosya6 = SunucuYeri6 & File6.FileName  
	File6.SaveAs YuklenecekDosya6 
	YuklenenDosyaninIsmi6 = File6.FileName
	
	Response.write "<script>alert('Dosya yünlendi.'); window.parent.location.href=('"& Request.Servervariables("HTTP_REFERER") &"');window.refresh;</script>"
	Response.End
	
end select
%>

<form ENCTYPE="multipart/form-data" action="default.asp?process=yukle" method="post">
<input type="file" name="dosya" />
<input type="submit" value="Yükle" />
</form>

</body>
</html>
