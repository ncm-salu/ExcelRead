package excel.read;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 


public class ExcelRead {
	public static XSSFWorkbook w;



	public static XSSFSheet s;



	public static FileInputStream f;


//return string , i raw , column j 
	public static String readStringData(int i,int j) throws IOException { //chance of exception file not found


//excel location
		
	f= new FileInputStream("C:\\Users\\user\\git\\ExcelRead\\ExcelRead\\src\\main\\resources\\Student.xlsx");


//file input stream object 
//XSSFWorkbook(f) to hold excel
	w= new XSSFWorkbook(f);


//get sheet 1 in s from w
	s= w.getSheet("Sheet1");
//get i th row from s to r
	Row r=s.getRow(i);


//get j th cell to c
	Cell c=r.getCell(j);


//to get string value(read string value getStringCellValue () method)
	return c.getStringCellValue();



	}


//to read integer data and convert to string 
	public static String readIntegerData(int i,int j) throws IOException {

		



			f= new FileInputStream("C:\\Users\\user\\git\\ExcelRead\\ExcelRead\\src\\main\\resources\\Student.xlsx");



			w= new XSSFWorkbook(f);



			s= w.getSheet("Sheet1");



			Row r=s.getRow(i);



			Cell c=r.getCell(j);


//numeric values like int float double etc using getNumericCellValue() 
//here given typecast to get int value
			int value=(int) c.getNumericCellValue();


//here return type is string so typecaste to string 
			return String.valueOf(value);



			}
}
