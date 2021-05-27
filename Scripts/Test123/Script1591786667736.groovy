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
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable
import org.openqa.selenium.Keys as Keys

WebUI.openBrowser('')

WebUI.navigateToUrl('http://newtours.demoaut.com/')

WebUI.maximizeWindow()

WebUI.waitForPageLoad(10)

WebUI.setText(findTestObject('Object Repository/Page_Welcome Mercury Tours/input_User                     Name_userName'), 
    variable)

WebUI.setText(findTestObject('Object Repository/Page_Welcome Mercury Tours/input_Password_password'), variable_0)

WebUI.delay(2)

String valueexcel = variable_1

System.out.println('row no is ' + valueexcel)

WebUI.click(findTestObject('Object Repository/Page_Welcome Mercury Tours/input_Password_login'))

WebUI.click(findTestObject('Object Repository/Page_Sign-on Mercury Tours/b_Welcome back to         Mercury Tours'))

String getdata = WebUI.getText(findTestObject('Page_Sign-on Mercury Tours/b_Welcome back to         Mercury Tours'))

System.out.println(getdata)

CustomKeywords.'exceldata.writeexcel.demoKey'(getdata, valueexcel)

WebUI.closeBrowser()

