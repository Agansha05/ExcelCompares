package org.excel;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	public static WebDriver driver;

	public static WebDriver edgeBrowser() {
		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
		return driver;
	}
	
	public static void chromeLaunch() {
		WebDriverManager.chromedriver().setup();
		driver=new ChromeDriver();
	}

	public static void browserLaunch(String name) {
		// if(name.equalsIgnoreCase("chrome"))
		WebDriverManager.edgedriver().setup();
		driver = new EdgeDriver();
	}

	public static void urlLaunch(String url) {
		driver.get(url);

	}

	public static void implicitlyWait(int sec) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(sec));
		driver.manage().window().maximize();
	}

	public static void sendkeys(WebElement e, String data) {
		e.sendKeys(data);
	}

	public static void click(WebElement e) {
		e.click();
	}

	// GET CURRENT URL
	public static String getCurrentUrl() {
		String u = driver.getCurrentUrl();
		return u;

	}

	// GET TITLE
	public static String getTitle() {
		String s = driver.getTitle();
		return s;
	}

	public static void quit() {
		driver.quit();
	}

	public static String getText(WebElement e) {
		String s = e.getText();
		return s;
	}

	public static String getAttribute(WebElement e) {
		String x = e.getAttribute("value");
		return x;
	}

	public static void moveToElement(WebElement target) {
		Actions a = new Actions(driver);
		a.moveToElement(target).perform();
	}

	public static void dragAdDrop(WebElement from, WebElement to) {
		Actions a = new Actions(driver);
		a.dragAndDrop(from, to).perform();

	}

	public static void actclick(WebElement c) {
		Actions a = new Actions(driver);
		a.click().perform();
	}

	public static void simpleAlert() {
		Alert d = driver.switchTo().alert();
		d.accept();
	}

	public void robot() throws AWTException {
		Robot r = new Robot();

	}

	public static void refresh() {
		driver.navigate().refresh();

	}

	public static void SelectByIndex(WebElement e, int index) {
		Select s = new Select(e);
		s.selectByIndex(index);
	}

	public static WebElement findElement(String loc, String value) {
		WebElement t = null;
		if (loc.equals("id")) {
			t = driver.findElement(By.id(value));
		} else if (loc.equals("name")) {
			t = driver.findElement(By.name(value));
		} else if (loc.equals("xpath")) {
			t = driver.findElement(By.xpath(value));
		}
		return t;
	}

	public static String readExcel(String filename, String sheet, int row, int c) throws IOException {
		File f = new File(System.getProperty("user.dir") + "\\src\\test\\resources\\Excel\\" + filename + ".xlsx");
		FileInputStream st = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(st);
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(row);
		Cell cell = r.getCell(c);
		int type = cell.getCellType();
		String value = null;
		if (type == 1) {
			value = cell.getStringCellValue();
		} else {
			if (DateUtil.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat si = new SimpleDateFormat();
				value = si.format(dateCellValue);
			} else {
				double numericCellValue = cell.getNumericCellValue();
				long num = (long) numericCellValue;
				value = String.valueOf(num);
			}
		}
		return value;
	}

	public static String writeExcel(String filename, String sheet, int row, int c, String Data) throws IOException {
		File f = new File(System.getProperty("user.dir") + "\\src\\test\\resources\\Excel\\" + filename + ".xlsx");
		FileInputStream st = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(st);
		Sheet s = w.getSheet(sheet);
		Row r = s.getRow(row);
		Cell createCell = r.createCell(c);
		createCell.setCellValue(Data);
		FileOutputStream fo = new FileOutputStream(f);
		w.write(fo);
		return Data;
	}

}
