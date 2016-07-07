<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>删除资料至Excel档案</title>
</head>
<body>
<center>
<h2>删除Excel表中的数据</h2>
<%
//设定FileInputStream读取Excel中的数据
FileInputStream finput=new FileInputStream(application.getRealPath("/")+"book1.xls");
POIFSFileSystem fs=new POIFSFileSystem(finput);
HSSFWorkbook wb=new HSSFWorkbook(fs);
//读取第一个工作表，其为sheet
HSSFSheet sheet=wb.getSheetAt(0);
finput.close();
//声明一列
HSSFRow row=null;
//声明一个存储格
HSSFCell cell=null;
//取出第3列
row=sheet.getRow((short)2);
if(row!=null)
	sheet.removeRow(row);
try{
	FileOutputStream fout=new FileOutputStream(application.getRealPath("/")+"book1.xls");
	wb.write(fout);
	fout.close();
	out.println("删除成功<a href='book1.xls'>book1.xls</a>");
}catch(IOException e)
{
	out.println("产生错误，错误信息："+e.toString());
}
%>
</center>
</body>
</html>