<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>插入资料至Excel档案</title>
</head>
<body>
<center>
<h2>向Excel中添加记录</h2>
<%
//设定FileInputStream读取Excel中的数据
FileInputStream finput=new FileInputStream("D:\\Jee\\dainchiimage\\book1.xls");
POIFSFileSystem fs=new POIFSFileSystem(finput);
HSSFWorkbook wb=new HSSFWorkbook(fs);
//读取第一个工作表，其为sheet
HSSFSheet sheet=wb.getSheetAt(0);
finput.close();
//声明一列
HSSFRow row=null;
//声明一个存储格
HSSFCell cell=null;
short i=4;
//建立一个新的列，注意是第五列（列及存储格都是从0开始）
row=sheet.createRow(i);
cell=row.createCell((short)0);
cell.setCellValue("UML");
cell=row.createCell((short)1);
cell.setCellValue("40");
cell=row.createCell((short)2);
cell.setCellValue("3");
cell=row.createCell((short)3);
//设定这个存储格为公式存储格，并输入公式
cell.setCellFormula("B"+(i+1)+"*C"+(i+1));
try{
	FileOutputStream fout=new FileOutputStream(application.getRealPath("/")+"book1.xls");
	wb.write(fout);
	fout.close();
	wb.close();
	out.println("存储成功<a href='book1.xls'>book1.xls</a>");
}catch(IOException e)
{
	out.println("产生错误，错误信息："+e.toString());
}
%>
</center>
</body>
</html>