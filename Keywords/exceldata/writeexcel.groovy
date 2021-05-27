package exceldata

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Row
import java.io.FileInputStream
import java.io.FileNotFoundException
import java.io.FileOutputStream



import internal.GlobalVariable

public class writeexcel {

	//private  static int i=1++;


	@Keyword
	public void demoKey(String TestResult, String valueexcel) throws IOException{
		FileInputStream file = new FileInputStream (new File("C://Users//Owner//Katalon Studio//Test//ExcelData//Book1.xlsx"))
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		XSSFSheet sheet = workbook.getSheetAt(0);
		//System.out.println(i)

		'Read data from excel'
		//String Data_fromCell=sheet.getRow(1).getCell(0).getStringCellValue();
		//int Data_fromCell=sheet.getRow(i).getCell(0).getColumnIndex();
		
				
			
		int i = Integer.parseInt(valueexcel);
		
		System.out.println('i value is '+i)
		
		'Write data to excel'
		
		sheet.getRow(i).createCell(3).setCellValue(TestResult);
		
		
		file.close();
		FileOutputStream outFile =new FileOutputStream(new File("C://Users//Owner//Katalon Studio//Test//ExcelData//Book1.xlsx"));
		workbook.write(outFile);
		outFile.close();
	}
}
