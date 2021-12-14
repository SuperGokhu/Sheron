package org.tes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Testin {
	public static void main(String[] args) throws IOException {
		File f = new File("D:\\Learning Materials\\eclipse\\eclipse-workspace\\MavenProject\\Excell\\Book1.xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet sh = w.getSheet("Sheet1");
		Row r = sh.getRow(1);
		Cell c = r.getCell(1);
		int cellType = c.getCellType();
		System.out.println(cellType);
		double numericCellValue = c.getNumericCellValue();
		long l = (long)numericCellValue;
		System.out.println(l);
	
	}

}
