package com.task;

import java.io.File;
import java.io.FileInputStream;

import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class App {

	public static void main(String[] args) throws IOException {

		
		FileInputStream fi = new FileInputStream(new File("nikunj.xls"));
		
		HSSFWorkbook wb = new HSSFWorkbook(fi);
		
		HSSFSheet sheet = wb.getSheetAt(0);
		
		
		FormulaEvaluator fe = wb.getCreationHelper().createFormulaEvaluator();
		
		for(Row row: sheet)
		{ 
			for(Cell cell :row )
			{
                switch(fe.evaluateInCell(cell).getCellType())
                {
                 case Cell.CELL_TYPE_NUMERIC:
                	 Double i = cell.getNumericCellValue();
                	 Object obj1 = i;
                	 System.out.print(obj1+"\t\t");
                	 
                	 //System.out.print(cell.getNumericCellValue()+"\t\t");
                	 break;
                 case Cell.CELL_TYPE_STRING:
                	 String s= cell.getStringCellValue();
                	 Object obj2 = s;
                	 System.out.print(obj2 + "\t\t");
                	 //System.out.print(cell.getStringCellValue()+"\t\t");	 
                     break;
                }
			}
			System.out.println();
		}
		
		
		 
		
		
		
		
	}

}
