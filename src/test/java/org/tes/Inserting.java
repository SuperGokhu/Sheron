package org.tes;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Inserting {
	public static void main(String[] args) throws IOException {
		File f = new File("D:\\\\Learning Materials\\\\eclipse\\\\eclipse-workspace\\\\MavenProject\\\\Excell\\\\Book2.xlsx");
		Workbook w = new XSSFWorkbook();
		Sheet cs = w.createSheet("hahaoya");
		Row createRow = cs.createRow(0);
		Cell createCell = createRow.createCell(0);
		createCell.setCellValue("hehawya");
		Cell createCell2 = createRow.createCell(1);
		createCell2.setCellValue("1234");
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
	}
	

}
