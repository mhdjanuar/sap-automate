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
SapKeywords.setText("wnd[0]/tbar[0]/okcd", "ME21N")
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Open Header
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB1:SAPLMEVIEWS:1100/subSUB1:SAPLMEVIEWS:4000/btnDYN_4000-BUTTON").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Input Vendor
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").toDispatch());
obj.setProperty("text", "37132");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Field org data
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").toDispatch());
obj.setProperty("text", "SMMG");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKGRP").toDispatch());
obj.setProperty("text", "01A");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").toDispatch());
obj.setProperty("text", "0010");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-BUKRS").toDispatch());
obj.setProperty("caretPosition", 3);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

//Isi Table
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,0]").toDispatch());
obj.setProperty("text", "1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-KNTTP[2,0]").toDispatch());
obj.setProperty("text", "k");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-TXZ01[5,0]").toDispatch());
obj.setProperty("text", "test");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").toDispatch());
obj.setProperty("text", "1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-MEINS[7,0]").toDispatch());
obj.setProperty("text", "PU");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-NETPR[10,0]").toDispatch());
obj.setProperty("text", "100");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-BPRME[13,0]").toDispatch());
obj.setProperty("text", "PU");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-WGBEZ[14,0]").toDispatch());
obj.setProperty("text", "z0001");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[15,0]").toDispatch());
obj.setProperty("text", "0010");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-WAERS[11,0]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-WAERS[11,0]").toDispatch());
obj.setProperty("caretPosition", 3);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")


/////////////////
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-KOSTL[5,0]").toDispatch());
obj.setProperty("text", "ANPB111");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-SAKTO[6,0]").toDispatch());
obj.setProperty("text", "5030201030");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-SAKTO[6,0]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:1303/tabsITEM_DETAIL/tabpTABIDT14/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-SAKTO[6,0]").toDispatch());
obj.setProperty("caretPosition", 10);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

//Klik verify
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[39]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");

//Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-VAROPTION1").toDispatch());
obj.invoke("press");


obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/sbar").toDispatch());
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Retrieve the text property
String text = obj.getProperty("text")

// Check if a match is found	
String poNumber = sapHelpers.getNumber(text)

if(poNumber) {
	GlobalVariable.poNumber = poNumber
	
	println("PO Number " + poNumber)
}

//Exit
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//===========================Create GR Attachment Maintaince ==================================
//Goto ZMM00027
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "ZMM00027");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

//Choose Create Attch
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/radR_CREATE").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/radR_CREATE").toDispatch());
obj.invoke("select");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

//Execute
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[8]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Klik new line and input 
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[14]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").toDispatch());
obj.setProperty("text", "Testing Lagi");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/sub:SAPLSPO4:0300/txtSVALD-VALUE[0,21]").toDispatch());
obj.setProperty("caretPosition", 13);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Klik Checklist
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Find File
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").toDispatch());
obj.setProperty("currentCellColumn", "FILNM");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").toDispatch());
obj.setProperty("firstVisibleColumn", "ACTIO");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").toDispatch());
obj.invoke("clickCurrentCell");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/ctxtDY_FILENAME").toDispatch());
obj.setProperty("text", "XL_logo_2016.svg.webp");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/ctxtDY_FILENAME").toDispatch());
obj.setProperty("caretPosition", 21);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Save attch
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[13]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Back
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//============================  Progres GR ==========================================

//Goto ML81N
SapKeywords.setText("wnd[0]/tbar[0]/okcd", "ML81N")
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Klik Orther Purchase Orther
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[17]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Input Purhase order
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/ctxtRM11R-EBELN").toDispatch());
obj.setProperty("text", "8099002360");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/ctxtRM11R-EBELN").toDispatch());
obj.setProperty("caretPosition", 3);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

//Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/mbar/menu[0]/menu[3]/menu[1]").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/txtRM11R-PLNPROZ").toDispatch());
obj.setProperty("text", "1");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/txtRM11R-PLNPROZ").toDispatch());
obj.setProperty("caretPosition", 1);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");


obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAB_HEADER/tabpESCR").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/txtESSR-TXZ01").toDispatch());
obj.setProperty("text", "Test");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAB_HEADER/tabpESCR/ssubSUBUSCR:SAPLXMLU:0399/ctxtW_GRFILES-ATTID").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAB_HEADER/tabpESCR/ssubSUBUSCR:SAPLXMLU:0399/ctxtW_GRFILES-ATTID").toDispatch());
obj.setProperty("caretPosition", 0);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/lbl[1,9]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/lbl[1,9]").toDispatch());
obj.setProperty("caretPosition", 10);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAB_HEADER/tabpESCR/ssubSUBUSCR:SAPLXMLU:0399/chkD_CHKBOX").toDispatch());
obj.setProperty("selected", true);
WebUI.delay(5)
sapHelpers.takeScreenshootWindows("20240212_0904")

// Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-OPTION1").toDispatch());
obj.invoke("press");
sapHelpers.takeScreenshootWindows("20240212_0904")






