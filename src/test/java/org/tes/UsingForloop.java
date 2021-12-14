package org.tes;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UsingForloop {
	public static void main(String[] args) throws IOException {
		File f =new File("D:\\Learning Materials\\eclipse\\\\eclipse-workspace\\\\MavenProject\\\\Excell\\\\iphone.xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet sh = w.getSheet("iphone lists");
		for (int i = 0; i < sh.getPhysicalNumberOfRows(); i++) {
			Row r = sh.getRow(i);
			System.out.println();
			for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
				Cell c = r.getCell(j);
				int cellType = c.getCellType();
				if (cellType==1) {
					String stringCellValue = c.getStringCellValue();
					System.out.print(stringCellValue+"\t\t");
				}else if (DateUtil.isCellDateFormatted(c)) {
					Date dateCellValue = c.getDateCellValue();
					SimpleDateFormat sf = new SimpleDateFormat("dd/MM/yyyy");
					String format = sf.format(dateCellValue);
					System.out.print(format+"\t\t");
				}else {
					double numericCellValue = c.getNumericCellValue();
					long l = (long)numericCellValue;
					System.out.print(l+"\t\t");
				}
				
			}
		}
		
	}

}
