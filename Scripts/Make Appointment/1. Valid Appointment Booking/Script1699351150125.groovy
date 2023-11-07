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
nameTestData = 'Data Files/Make Appointment'

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

    facility = data.getValue('facility', excelRow)

    visit_date = data.getValue('visit_date', excelRow)

    comment = data.getValue('comment', excelRow)
}

WebUI.callTestCase(findTestCase('Login/1. Valid Login'), [:], FailureHandling.STOP_ON_FAILURE)

if (WebUI.verifyElementPresent(findTestObject('Make Appointment/select_facility'), 2)) {
    WebUI.selectOptionByValue(findTestObject('Make Appointment/select_facility'), facility, false)

    WebUI.delay(1)

    WebUI.check(findTestObject('Make Appointment/Check_Apply for hospital readmission'))

    WebUI.verifyElementChecked(findTestObject('Make Appointment/Check_Apply for hospital readmission'), 2)

    WebUI.click(findTestObject('Make Appointment/input_Medicaid'))

    WebUI.delay(1)

    WebUI.setText(findTestObject('Make Appointment/input_Visit Date'), visit_date)

    WebUI.delay(1)

    WebUI.delay(1)

    WebUI.setText(findTestObject('Make Appointment/textarea_Comment'), comment)

    WebUI.delay(2)

    WebUI.click(findTestObject('Make Appointment/button_Book Appointment'))

    WebUI.waitForElementPresent(findTestObject('Make Appointment/Text_Appointment Confirmation'), 2)

    WebUI.verifyElementPresent(findTestObject('Make Appointment/Text_Appointment Confirmation'), 2)
} else {
    WebUI.closeBrowser()
}

