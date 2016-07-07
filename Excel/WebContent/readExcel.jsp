<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>读取Excel档案</title>
</head>
<body>
<center>
<h2>Jsp读取excel中的数据</h2>
<table border="1" width="100%">
<%
//设定FileInputStream读取Excel中的数据
FileInputStream finput=new FileInputStream(application.getRealPath("/")+"book1.xls");
POIFSFileSystem fs=new POIFSFileSystem(finput);
HSSFWorkbook wb=new HSSFWorkbook(fs);
//读取第一个工作表，其为sheet
HSSFSheet sheet=wb.getSheetAt(0);
wb.close();
finput.close();
//声明一列
HSSFRow row=null;
//声明一个存储格
HSSFCell cell=null;
short i=0;
short y=0;
//读取所有存储格资料
for(i=0;i<=sheet.getLastRowNum();i++)
{ 
	out.println("<tr>");
	row=sheet.getRow(i);
	for(y=0;y<row.getLastCellNum();y++)
	{
		cell=row.getCell(y);
		out.println("<td>");
		//判断存储格的格式
		switch(cell.getCellType())
		{
		     case HSSFCell.CELL_TYPE_NUMERIC:
		    	 out.println(cell.getNumericCellValue());
		    	 break;
		     case HSSFCell.CELL_TYPE_STRING:
		    	 out.println(cell.getStringCellValue());
		    	 break;
		     case HSSFCell.CELL_TYPE_FORMULA:
		    	 out.println(cell.getNumericCellValue());
		    	 break;
		     default:
		    	 out.println("不明格式");
		    	 break;
		}
	}
	out.println("</td>");
}
out.println("</tr>");
%>
</table>
</center>
</body>
</html>