package utilis;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Xlutils {


    public static FileInputStream file;
	public static FileOutputStream fo;
	public static Workbook wb;
	public static Sheet st;
	public static Row row;
	public static Cell cell;
	public static CellStyle style;

	public  static int getRowCount(String xlfile,String xlsheet) throws IOException  {

		file=new FileInputStream(xlfile) ;
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		int rowcount=st.getLastRowNum();
		wb.close();
		return rowcount;


	}


	public static short getcolumncount(String xlfile,String xlsheet,int rownum) throws IOException {

		file=new FileInputStream(xlfile) ;
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		short colcount=row.getLastCellNum();
		wb.close();

		return colcount;
	}

	public static String getStringCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException {

		file=new FileInputStream(xlfile) ;
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);

		String data;
		try {
			cell=row.getCell(colnum);

			data=cell.getStringCellValue();

		} catch (Exception e) {
			data= " ";
		}
		wb.close();
		return data;
	}

	public static double getNumericCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException {

		file=new FileInputStream(xlfile) ;
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);

		double data;
		try {
			cell=row.getCell(colnum);

			data=cell.getNumericCellValue();

		} catch (Exception e) {
			data= 0.0;
		}
		wb.close();
		return data;
	}



	public static boolean getboolenCellData(String xlfile,String xlsheet,int rownum,int colnum) throws IOException {

		file=new FileInputStream(xlfile) ;
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);

		boolean data;
		try {
			cell=row.getCell(colnum);

			data=cell.getBooleanCellValue();

		} catch (Exception e) {
			data= false;
		}
		wb.close();
		return data;
	}

	public static void setCellData(String xlfile,String xlsheet,int rownum,int colnum,String data) throws IOException
	{
		file=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.createCell(colnum);
		cell.setCellValue(data);

		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();



	}


	public static void setCellData(String xlfile,String xlsheet,int rownum,int colnum,double data) throws IOException
	{
		file=new FileInputStream(xlfile);
		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.createCell(colnum);
		cell.setCellValue(data);

		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();



	}


	public static void fillgreenColor(String xlfile,String xlsheet,int rownum,int colnum) throws IOException {

		file=new FileInputStream(xlfile);

		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.getCell(colnum);

		style=wb.createCellStyle();

		style.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());	
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell.setCellStyle(style);
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();

	}


	public static void fillredColor(String xlfile,String xlsheet,int rownum,int colnum) throws IOException {

		file=new FileInputStream(xlfile);

		wb=new XSSFWorkbook(file);
		st=wb.getSheet(xlsheet);
		row=st.getRow(rownum);
		cell=row.getCell(colnum);

		style=wb.createCellStyle();

		style.setFillForegroundColor(IndexedColors.RED.getIndex());	
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		cell.setCellStyle(style);
		fo=new FileOutputStream(xlfile);
		wb.write(fo);
		wb.close();

	}




}



