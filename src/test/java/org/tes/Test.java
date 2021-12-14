package org.tes;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	public static void main(String[] args) throws IOException {
		File f = new File("D:\\Learning Materials\\eclipse\\eclipse-workspace\\MavenProject\\Excell\\testin.xlsx");
	    FileInputStream fi = new FileInputStream(f);
	    Workbook w = new XSSFWorkbook(fi);
	    Sheet s = w.getSheet("Sheet1");
	    for (int i = 0; i < s.getPhysicalNumberOfRows(); i++) {
	    	Row r = s.getRow(i);
	    	for (int j = 0; j < r.getPhysicalNumberOfCells(); j++) {
	    		Cell ctype = r.getCell(j);
	    		int cellType = ctype.getCellType();
				if (cellType==1) {
					String stringCellValue = ctype.getStringCellValue();
					System.out.print(stringCellValue+"\t\t\t");
				}else {
					if ( (DateUtil.isCellDateFormatted(ctype)) ) {
		    			Date dateCellValue = ctype.getDateCellValue();
		    			SimpleDateFormat sf = new SimpleDateFormat("MM/dd/yyyy");
		    			String format = sf.format(dateCellValue);
		    			System.out.print(format);
		    			System.out.println();
				}
			}
	    	}
	    }
	}
}
