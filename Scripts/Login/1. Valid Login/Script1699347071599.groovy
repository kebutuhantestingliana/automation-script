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
import org.openqa.selenium.Keys as Keys

'Mapping Excel File'
nameTestData = 'Login'

GlobalVariable.noRow = 0

TestData data = findTestData(nameTestData)

data.changeSheet('Sheet1')

getLastRow = data.getRowNumbers()

GlobalVariable.getLastRow = getLastRow

for (int excelRow : (1..getLastRow)) {
    if (data.getValue('No', excelRow) == '') {
        break
    }
    
    'Import Sheet Sign in'
    GlobalVariable.captureCount = 0

    GlobalVariable.noRow = (GlobalVariable.noRow + 1)

    GlobalVariable.a = 0

    GlobalVariable.excelRow = excelRow

    username = data.getValue('username', excelRow)

    password = data.getValue('password', excelRow)
}

'open browser'
WebUI.openBrowser('')

'navigate to url'
WebUI.navigateToUrl(GlobalVariable.url)

WebUI.maximizeWindow(FailureHandling.STOP_ON_FAILURE)

if (WebUI.verifyElementPresent(findTestObject('Login/CURA Healthcare Service'), 2)) {
    WebUI.waitForElementClickable(findTestObject('Login/a_menu'), 2)

    WebUI.click(findTestObject('Login/a_menu'))

    WebUI.waitForElementPresent(findTestObject('Login/a_Login'), 2)

    WebUI.click(findTestObject('Login/a_Login'))

    WebUI.waitForElementPresent(findTestObject('Login/Text_Login'), 0)

    WebUI.setText(findTestObject('Login/input_Username'), username)

    WebUI.delay(1)

    WebUI.setText(findTestObject('Login/input_Password'), password)

    WebUI.delay(1)

    WebUI.click(findTestObject('Login/button_Login'))

    WebUI.verifyElementPresent(findTestObject('Make Appointment/Text_Make Appointment'), 3)

    WebUI.delay(2)
} else {
    WebUI.closeBrowser()
}

