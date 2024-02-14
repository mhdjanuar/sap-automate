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

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
arg = new Variant[3];
arg[0] = new Variant(100);
arg[1] = new Variant(29);
arg[2] = new Variant(false);
obj.invoke("resizeWorkingPane", arg);

//Goto VA01
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VA01");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindow("20240212_0904")

// Insert Order Type and Sales Organization
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-AUART").toDispatch());
obj.setProperty("text", "ZO");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VKORG").toDispatch());
obj.setProperty("text", "EXCL");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VTWEG").toDispatch());
obj.setProperty("text", "E1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-SPART").toDispatch());
obj.setProperty("text", "E1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-SPART").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-SPART").toDispatch());
obj.setProperty("caretPosition", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);


//Insert Header 
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").toDispatch());
obj.setProperty("text", "stepbystep");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").toDispatch());
obj.setProperty("text", "09.02.2024");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").toDispatch());
obj.setProperty("text", "PAYOVOJKT");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").toDispatch());
obj.setProperty("text", "PAYOVOJKT");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").toDispatch());
obj.setProperty("caretPosition", 10);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);


// Add First Item
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.setProperty("text", "OVOAX1000K");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.setProperty("caretPosition", 11);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");

// Edit Quality
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);


obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/txtRV45A-KWMENG").toDispatch());
obj.setProperty("text", "100");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);


obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[3]").toDispatch());
obj.invoke("press");

// Klik Material
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);

// Go to sales B
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\02").toDispatch());
obj.invoke("select");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4458/ctxtVBKD-BZIRK").toDispatch());
obj.setProperty("text", "SD02");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\02/ssubSUBSCREEN_BODY:SAPMV45A:4458/ctxtVBKD-BZIRK").toDispatch());
obj.invoke("setFocus");

//Goto AccountAssigment
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4457/subCOBL:SAPLKACB:1006/ctxtCOBL-PRCTR").toDispatch());
obj.setProperty("text", "COR");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4457/subCOBL:SAPLKACB:1006/ctxtCOBL-PRCTR").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4457/subCOBL:SAPLKACB:1006/ctxtCOBL-PRCTR").toDispatch());
obj.setProperty("caretPosition", 3);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[3]").toDispatch());
obj.invoke("press");

// Save and store the orderNumber to global
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/sbar").toDispatch());

// Retrieve the text property
String text = obj.getProperty("text")

// Check if a match is found
String orderNumber = sapHelpers.getNumber(text)

if(orderNumber) {
	GlobalVariable.orderNumber = orderNumber
}

// Go to Menu Again
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//Goto VA02
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VA02");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Goto Detail
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VBELN").toDispatch());
obj.setProperty("caretPosition", 10);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Remove Item
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[14]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-OPTION1").toDispatch());
obj.invoke("press");

//Add New Item
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.setProperty("text", "OVOAX1000K");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.setProperty("caretPosition", 11);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");


obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);

//Edit Quality
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/txtRV45A-KWMENG").toDispatch());
obj.setProperty("text", "3");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Goto AccountAssigment Edit Profil Center
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06/ssubSUBSCREEN_BODY:SAPMV45A:4457/subCOBL:SAPLKACB:1006/ctxtCOBL-PRCTR").toDispatch());
obj.setProperty("text", "XL001");

////Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//Close Application
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("close");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-OPTION1").toDispatch());
obj.invoke("press");

