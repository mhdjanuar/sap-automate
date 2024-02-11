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

import com.jacob.activeX.ActiveXComponent
import com.jacob.com.Variant
import com.jacob.com.Dispatch
import com.katalon.sap.SapKeywords
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.xlaxiata.helper.SapHelpers

WebUI.callTestCase(findTestCase('Common/StartSapLogon'), [:], FailureHandling.STOP_ON_FAILURE)

WebUI.delay(5)

WebUI.callTestCase(findTestCase('Common/Login'), [:], FailureHandling.STOP_ON_FAILURE)

ActiveXComponent obj;
Variant[] arg;
def session = SapKeywords.getSession()

SapHelpers sapHelpers = new SapHelpers()

//Goto VA01
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "ME21N");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

// Close sidebar
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/shellcont/shell/shellcont[0]/shell").toDispatch());
//obj.invoke("pressButton", "HIDEHELP");

//// Open Header
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").toDispatch());
//obj.invoke("press");
//
//// Tab Org Data
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8").toDispatch());
//obj.invoke("select");
//
//// Field org data
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").toDispatch());
//obj.setProperty("text", "SMMG");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").toDispatch());
//obj.setProperty("text", "01A");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").toDispatch());
//obj.setProperty("text", "0010");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").toDispatch());
//obj.invoke("setFocus");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").toDispatch());
//obj.setProperty("caretPosition", 3);
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
//obj.invoke("sendVKey", 0);


//WebUI.delay(5)

// Open Item Overview
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4001/btnDYN_4000-BUTTON").toDispatch());
//obj.invoke("press");

// Equivalent Java code
oGridView = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").toDispatch());
String sArrivalCityColumnName = Dispatch.call(oGridView.getProperty("Columns").toDispatch(), "Item", 1);
