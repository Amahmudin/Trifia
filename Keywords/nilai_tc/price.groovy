package nilai_tc

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
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords

import internal.GlobalVariable



public class price {
	@Keyword
	def TradeConfirmationPrice() {
		// enter your code here
		// you can use either Groovy or Java
		
		//Define Dataset File
		def workBook = ExcelKeywords.getWorkbook((GlobalVariable.dirProject + '\\Data Files\\') + GlobalVariable.formExcel)
		
		//Choose sheet page
		def sheet1 = ExcelKeywords.getExcelSheet(workBook, 'Data')
		
		//Choose Column/Row
		int price = ExcelKeywords.getCellValueByAddress(sheet1, 'B2')

		//add logic
		if (price <=80) {
			price = 80
		}
		return price
	}
}
