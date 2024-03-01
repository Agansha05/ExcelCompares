package org.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSheet extends BaseClass{
	public static void main(String[] args) throws IOException {
//	    File f = new File("C:\\Users\\agans\\eclipse-workspace\\ExcelCompares\\src\\test\\resources\\folder\\Excelone.xlsx");
//	    FileInputStream fi = new FileInputStream(f);
//		Workbook w = new XSSFWorkbook(fi);
//		Sheet sh = w.getSheet("sheet1");
//		Row r = sh.getRow(0);
//		Cell c = r.getCell(0);
//		String s = c.getStringCellValue();
		List l1= new LinkedList();
		List l2= new LinkedList();
		for (int i =0;i<5;i++) {
			for(int j=0;j<1;j++) {
				System.out.println(readExcel("Excelone","Sheet1",i,j));
				System.out.println(readExcel("Excel2","Sheet1",i,j));
//				l1.add(readExcel("Excelone","Sheet1",i,j));
//				l2.add(readExcel("Excel2","Sheet1",i,j));
				
			}
		}
		System.out.println(l1.get(0).equals(l2.get(1)));
		
		
			
	}
	}
   

	


