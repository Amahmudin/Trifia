import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import com.kms.katalon.keyword.excel.ExcelKeywords as ExcelKeywords
import org.openqa.selenium.Keys as Keys
import org.apache.commons.lang3.StringUtils

/*WebUI.openBrowser('')

WebUI.navigateToUrl('https://172.18.139.99:8080/interface/')

WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE/input_GUAVA INTERFACE_username'), 'YUSRAN')

WebUI.setEncryptedText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE/input_GUAVA INTERFACE_password'), 
   'aeHFOx8jV/A=')

WebUI.click(findTestObject('Create TC RCL/Page_GUAVA INTERFACE/button_Login'))
*/
WebUI.click(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/a_Trade Confirmation'))

WebUI.click(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/a_Create TC'))

WebUI.selectOptionByValue(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/select_Select Template TC                  _793bc1'), 
    'RCL', true)

WebUI.selectOptionByValue(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/select_Select Side                         _297bf7'), 
    'B', true)

WebUI.selectOptionByValue(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/select_Select Counterparty Name            _216b08'), 
    'CIJA', true)

WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__emailCounterParty'), 
    'a.mahmudin15@gmail.com')

WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__isin'), 'IDG000018607 -- FR0090')

WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__deal_time'), '05-05-2023 10:52:24')

//Devine Today Date
Date todaysDate = new Date();
def FormatDate = todaysDate.plus(0).format("MM/dd/yyyy");
WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__trade_date'), FormatDate)

WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__value_date'), FormatDate)
WebUI.delay(1)

WebUI.click(findTestObject('Object Repository/Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/div_Template'))
WebUI.delay(1)

//-------------------Get Data from Excel ------------------
//Define Dataset File
def workBook = ExcelKeywords.getWorkbook((GlobalVariable.dirProject + '\\Data Files\\') + GlobalVariable.formExcel)

//Choose sheet page
def sheet1 = ExcelKeywords.getExcelSheet(workBook, 'Data')

//Choose Column/Row
int quantity = ExcelKeywords.getCellValueByAddress(sheet1, 'A2')
//-------------------------End------------------------------

'Input Quantity'
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__quantity'), Keys.chord(Keys.CONTROL, 'a'))	
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__quantity'), Keys.chord(quantity.toString()))

'Input Price'
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__price'), Keys.chord(Keys.CONTROL, 'a'))
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__price'), Keys.chord(CustomKeywords.'nilai_tc.price.TradeConfirmationPrice'().toString()))

'Input Interest'
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__acc_interest'), Keys.chord(Keys.CONTROL, 'a'))
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__acc_interest'), Keys.chord(CustomKeywords.'nilai_tc.interest.TradeConfirmationInterest'().toString()))

'Input Tax Gain'
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__capital_gain_tax'), Keys.chord(Keys.CONTROL, 'a'))
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__capital_gain_tax'), Keys.chord(CustomKeywords.'nilai_tc.taxGain.TradeConfirmationTaxGain'().toString()))

'Input Tax Interest'
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__tax_on_interest'), Keys.chord(Keys.CONTROL, 'a'))
WebUI.sendKeys(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input__tax_on_interest'), Keys.chord(CustomKeywords.'nilai_tc.taxInterest.TradeConfirmationTaxInterest'().toString()))

WebUI.selectOptionByValue(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/select_Select Settlement Instruction       _60ee55'), 
    'dvp', true)

WebUI.setText(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/textarea__additional_info'), 
    'Test Automation')

WebUI.click(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/button_Generate and Send TC'))



WebUI.delay(3)
'Berhasil Create TC menggunakan template RCL side Buy'
WebUI.takeScreenshot()

WebUI.click(findTestObject('Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/button_Yes'))

WebUI.takeScreenshot()

refnum = WebUI.getText(findTestObject('Object Repository/Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/Trade Confirmation Result'))

String refnumText = refnum
String newString = StringUtils.substringBetween(refnum,"Number "," Delivery")
//String excelFilePath = 'Data Set/' + GlobalVariable.formExcel

//String sheetName = 'Fund MT200'

//
//
//workbook01 = ExcelKeywords.getWorkbook(excelFilePath)
//
//sheet01 = ExcelKeywords.getExcelSheet(workbook01, sheetName)
//
//ExcelKeywords.setValueToCellByIndex(sheet01, 1, 14, refnum)

//ExcelKeywords.setValueToCellByIndex(sheet01, 1, 13, refnum)
//ExcelKeywords.saveWorkbook(excelFilePath, workbook01)

//WebUI.comment(refnumText)
//WebUI.comment(newString)

WebUI.setText(findTestObject('Object Repository/Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/input_Search_form-control input-sm'),'78')
WebUI.click(findTestObject('Object Repository/Create TC RCL/Page_GUAVA INTERFACE version 1.4.3/a_Detail_After_Create'))

WebUI.takeFullPageScreenshot()