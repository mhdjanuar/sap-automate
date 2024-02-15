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
import com.xlaxiata.pageobjects.web.CreateMaterial

CreateMaterial createMaterial = new CreateMaterial()

WebUI.openBrowser('')

WebUI.navigateToUrl('https://qas.sap.xl.co.id/sap/bc/ui2/flp')

WebUI.maximizeWindow()

WebUI.click(findTestObject('Web/BtnCancel'))

WebUI.delay(5)

createMaterial.login()

WebUI.delay(5)

// ========================== ZEQP – Asset XL – Equipment =======================
WebUI.click(findTestObject('Web/CardMaterialService'))
createMaterial.inputPersonalNo();

WebUI.click(findTestObject('Web/BtnSinggleReq'))
createMaterial.inputRequestNote("ZEQP")
createMaterial.selectWithArrowDown(findTestObject('Web/SelectMaterial'), 2)
createMaterial.selectWithArrowDown(findTestObject('Web/SelectBaseUoM'), 3)
createMaterial.zeqpAssetXLEquipment()

//================================
WebUI.click(findTestObject('Web/CardMaterialService'))
createMaterial.inputPersonalNo();

WebUI.click(findTestObject('Web/BtnSinggleReq'))
createMaterial.inputRequestNote("ZFGT")
createMaterial.selectWithArrowDown(findTestObject('Web/SelectMaterial'), 3)
createMaterial.selectWithArrowDown(findTestObject('Web/SelectBaseUoM'), 2)
createMaterial.zfgtNonAssetXL()

//================================
WebUI.click(findTestObject('Web/CardMaterialService'))
createMaterial.inputPersonalNo();

WebUI.click(findTestObject('Web/BtnSinggleReq'))
createMaterial.inputRequestNote("ZHAW")
createMaterial.selectWithArrowDown(findTestObject('Web/SelectMaterial'), 4)
createMaterial.inputMaterialNo()
createMaterial.selectWithArrowDown(findTestObject('Web/SelectBaseUoM'), 2)
createMaterial.zhawXLTradingGoods()

//================================
WebUI.click(findTestObject('Web/CardMaterialService'))
createMaterial.inputPersonalNo();

WebUI.click(findTestObject('Web/BtnSinggleReq'))
createMaterial.inputRequestNote("ZNWS")
createMaterial.selectWithArrowDown(findTestObject('Web/SelectMaterial'), 5)
createMaterial.selectWithArrowDown(findTestObject('Web/SelectBaseUoM'), 2)
createMaterial.znwsAssetXlNonEquip()

//================================
WebUI.click(findTestObject('Web/CardMaterialService'))
createMaterial.inputPersonalNo();

WebUI.click(findTestObject('Web/BtnSinggleReq'))
createMaterial.inputRequestNote("ZSER")
createMaterial.selectWithArrowDown(findTestObject('Web/SelectMaterial'), 6)
createMaterial.inputMaterialNo()
createMaterial.selectWithArrowDown(findTestObject('Web/SelectBaseUoM'), 2)
createMaterial.zserService()

//================================
WebUI.click(findTestObject('Web/CardMaterialService'))
createMaterial.inputPersonalNo();

WebUI.click(findTestObject('Web/BtnSinggleReq'))
createMaterial.inputRequestNote("ZSFW")
createMaterial.selectWithArrowDown(findTestObject('Web/SelectMaterial'), 7)
createMaterial.selectWithArrowDown(findTestObject('Web/SelectBaseUoM'), 2)
createMaterial.zsfwXlSoftware()


