package Xl;

import java.io.IOException;

public class Xlread {
	
	public static void main(String[] args) throws IOException   {
	int x=	Xlutils.getRowCount("C:\\Users\\DELL\\Desktop\\Employee.xlsx","Sheet1");
	System.out.println(x);
int y=	Xlutils.getcolumncount("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 1);
	System.out.println(y);
	
String ur=	Xlutils.getStringCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 1, 0);
String pw=	Xlutils.getStringCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 1, 1);
	System.out.println(ur+" "+pw);
	
double sal=	Xlutils.getNumericCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 2, 2);
	System.out.println(sal);

	
boolean val=	Xlutils.getboolenCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 3, 3);
System.out.println(val);	
	
Xlutils.setCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 1, 4, "pass"); 
Xlutils.setCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 2, 4, "Fail"); 


Xlutils.setCellData("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 1, 4, 100.5);
	
Xlutils.fillgreenColor("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 1, 4);
Xlutils.fillredColor("C:\\\\Users\\\\DELL\\\\Desktop\\\\Employee.xlsx", "Sheet1", 2, 4);
	}

}
