package org.tes;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;

public class AmazonToExcell {
	public static void main(String[] args) throws IOException {
		File f = new File("D:\\Learning Materials\\eclipse\\eclipse-workspace\\MavenProject\\Excell\\iphone.xlsx");
		Workbook w =new XSSFWorkbook();
		Sheet cs = w.createSheet("iphone lists");
		
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.amazon.in/");
		WebElement input = driver.findElement(By.id("twotabsearchtextbox"));
		input.sendKeys("iphone pro",Keys.ENTER);
		List<WebElement> allphones = driver.findElements(By.xpath("//span[@class='a-size-medium a-color-base a-text-normal']"));
		for (int i = 0; i < allphones.size(); i++) {
			WebElement wee = allphones.get(i);
			String text = wee.getText();
			System.out.println(text);
			Row row = cs.createRow(i);
			Cell cell = row.createCell(0);
			cell.setCellValue(text);
		}
		
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
	  
	}

}
