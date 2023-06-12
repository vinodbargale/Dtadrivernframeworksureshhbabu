import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Rough {

	public static void main(String[] args) throws Throwable {
		/*
		FileInputStream file=new FileInputStream("C:\\Users\\DELL\\Desktop\\Employee.xlsx");
		XSSFWorkbook wb=new XSSFWorkbook(file);
		XSSFSheet st=wb.getSheet("Sheet1");
		int row=st.getLastRowNum();
		System.out.println(row);
		
		
		 * short row1=st.getRow(1).getLastCellNum(); System.out.println(row1);
		 * 
		 * 
		 * 
		 * XSSFRow r1=st.getRow(0); XSSFRow r2=st.getRow(1);
		 * 
		 * short c1=r1.getLastCellNum(); short c2=r2.getLastCellNum();
		 * System.out.println(c1); System.out.println(c2);
		 
		
		//to read one cell data
		XSSFRow r1=st.getRow(1);
		


		
		
		
		 * XSSFCell c1=r1.getCell(0); XSSFCell c2=r1.getCell(1); XSSFCell
		 * c3=r1.getCell(2); XSSFCell c4=r1.getCell(3); String
		 * em=c1.getStringCellValue();
		 * 
		 * String em2=c2.getStringCellValue(); double em3=c3.getNumericCellValue();
		 * boolean em4=c4.getBooleanCellValue();
		 
		  
		 
		 
		int lastrow=st.getLastRowNum();
		
		for(int i=1;i<=lastrow;i++) {
			
			XSSFRow r5=st.getRow(i);
			
			XSSFCell c5=r5.getCell(0);
			XSSFCell c6=r5.getCell(1);XSSFCell c7=r5.getCell(2);XSSFCell c8=r5.getCell(3);
		String	em0=c5.getStringCellValue();
			  
			  String em5=c6.getStringCellValue(); double em6=c7.getNumericCellValue();
			  boolean em7=c8.getBooleanCellValue();
			  System.out.println(em0+" "+em5+" "+em6+" "+em7);
		}
		
		
		
	}*/
		
		
	FileInputStream file=new FileInputStream("C:\\\\Users\\\\DELL\\\\Desktop\\\\write1.xlsx");
	
	XSSFWorkbook wb=new XSSFWorkbook(file);
		
		XSSFSheet st=wb.getSheet("Sheet2");
		
		XSSFRow r1=st.getRow(1);
		
		
		XSSFCell c1=r1.createCell(2);
		c1.setCellValue("pass");
		
		
		CellStyle pass=wb.createCellStyle();
		pass.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		pass.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		c1.setCellStyle(pass);
		
		
		
		XSSFCell c2=r1.createCell(3);
		c2.setCellValue("false");
		
		CellStyle fail=wb.createCellStyle();
		fail.setFillForegroundColor(IndexedColors.RED.index);
		fail.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		c2.setCellStyle(fail);
		
		FileOutputStream out=new FileOutputStream("C:\\\\Users\\\\DELL\\\\Desktop\\\\write1.xlsx");
		
		wb.write(out);
		out.close();
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
