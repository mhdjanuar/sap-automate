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

//Goto VF01
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VF01");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

// Klik Document
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);

//Klik Filter
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[17]").toDispatch());
obj.invoke("press");

//Input Filter
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[0,24]").toDispatch());
obj.setProperty("text", GlobalVariable.orderNumber);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[3,24]").toDispatch());
obj.setProperty("text", "EXCL");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[4,24]").toDispatch());
obj.setProperty("text", "E1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]").toDispatch());
obj.setProperty("text", "E1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/ctxtG_SELFLD_TAB-LOW[5,24]").toDispatch());
obj.setProperty("caretPosition", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]").toDispatch());
obj.invoke("sendVKey", 0);

//Choose Document
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/lbl[1,3]").toDispatch());
obj.setProperty("caretPosition", 7);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]").toDispatch());
obj.invoke("sendVKey", 0);

////Save Document - this ambigu if i click save their not show in table
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
//obj.invoke("press");

//Exit
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//Goto VF04
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VF04");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Remove From Date
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtP_FKDAB").toDispatch());
obj.setProperty("text", "");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtP_FKDAB").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtP_FKDAB").toDispatch());
obj.setProperty("caretPosition", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Input SD Document
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/txtS_VBELN-LOW").toDispatch());
obj.setProperty("text", GlobalVariable.orderNumber);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/txtS_VBELN-LOW").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/txtS_VBELN-LOW").toDispatch());
obj.setProperty("caretPosition", 10);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Check Document to be Selected
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTABSTRIP_A_TS/tabpTSSE/ssub%_SUBSCREEN_A_TS:SDBILLDL:8001/chkP_ALLEA").toDispatch());
obj.setProperty("selected", true);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTABSTRIP_A_TS/tabpTSSE/ssub%_SUBSCREEN_A_TS:SDBILLDL:8001/txt%_P_SORT_%_APP_%-TEXT").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTABSTRIP_A_TS/tabpTSSE/ssub%_SUBSCREEN_A_TS:SDBILLDL:8001/txt%_P_SORT_%_APP_%-TEXT").toDispatch());
obj.setProperty("caretPosition", 14);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Go to Display Biling List
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[8]").toDispatch());
obj.invoke("press");

// Klik Individual Billinf Document
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[1]/btn[18]").toDispatch());
obj.invoke("press");

// Klik Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/sbar").toDispatch());

// Retrieve the text property
String text = obj.getProperty("text")

// Check if a match is found
String billingNumber = sapHelpers.getNumber(text)

if(billingNumber) {
	println("Billing Number : " + billingNumber)
	GlobalVariable.billingNumber = billingNumber
}

// Exit
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//Goto VF02
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VF02");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Input Billing Document
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBRK-VBELN").toDispatch());
obj.setProperty("text", GlobalVariable.billingNumber);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBRK-VBELN").toDispatch());
obj.setProperty("caretPosition", 8);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Klik Tombol Display Headers
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/btnTC_HEAD").toDispatch());
obj.invoke("press");

//Input Date In Headers Data
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTABSTRIP_OVERVIEW/tabpKFDE/ssubSUBSCREEN_BODY:SAPMV60A:6105/ctxtVBRK-FKDAT").toDispatch());
obj.setProperty("caretPosition", 9);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/cntlCONTAINER/shellcont/shell").toDispatch());
obj.setProperty("focusDate", "20240210");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/cntlCONTAINER/shellcont/shell").toDispatch());
obj.setProperty("selectionInterval", "20240210,20240210");

//Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");

//Exit
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

//=========================== Flow Contract ======================================

//Go to VA41
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VA41");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Input Create Contracts
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-AUART").toDispatch());
obj.setProperty("text", "ZN1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VKORG").toDispatch());
obj.setProperty("text", "BSXL");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VTWEG").toDispatch());
obj.setProperty("text", "20");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-SPART").toDispatch());
obj.setProperty("text", "B1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-SPART").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-SPART").toDispatch());
obj.setProperty("caretPosition", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Input Header Contract
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUAGV-KUNNR").toDispatch());
obj.setProperty("text", "746639");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/ctxtKUWEV-KUNNR").toDispatch());
obj.setProperty("text", "746639");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBKD-BSTDK").toDispatch());
obj.setProperty("caretPosition", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/cntlCONTAINER/shellcont/shell").toDispatch());
obj.setProperty("selectionInterval", "20240211,20240211");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Choose Material
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.setProperty("text", "LL-01");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/tabsG_SELONETABSTRIP/tabpTAB001/ssubSUBSCR_PRESEL:SAPLSDH4:0220/sub:SAPLSDH4:0220/txtG_SELFLD_TAB-LOW[2,24]").toDispatch());
obj.setProperty("caretPosition", 5);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]").toDispatch());
obj.invoke("sendVKey", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");

// Go to Detail Material
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\07").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnBT_PFPL").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");

//Input Target Quantity
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/txtVBAP-ZMENG").toDispatch());
obj.setProperty("text", "1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/txtVBAP-ZMENG").toDispatch());
obj.setProperty("caretPosition", 18);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Klik Tab Contract

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\03").toDispatch());
obj.invoke("select");

//Input Contract Data
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/cntlCONTAINER/shellcont/shell").toDispatch());
obj.setProperty("selectionInterval", "20240211,20240211");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/ctxtVEDA-VENDDAT").toDispatch());
obj.setProperty("text", "10.02.2025");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/cmbVEDA-VLAUFK").toDispatch());
obj.setProperty("key", "03");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/txtVEDA-VLAUFZ").toDispatch());
obj.setProperty("text", "12");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/cmbVEDA-VLAUEZ").toDispatch());
obj.setProperty("key", "3");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/ctxtVEDA-VABNDAT").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/ctxtVEDA-VABNDAT").toDispatch());
obj.setProperty("caretPosition", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/cntlCONTAINER/shellcont/shell").toDispatch());
obj.setProperty("selectionInterval", "20240211,20240211");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/cmbVEDA-VLAUFK").toDispatch());
obj.setProperty("key", "02");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/cmbVEDA-VLAUFK").toDispatch());
obj.setProperty("key", "01");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/txtVEDA-VLAUFZ").toDispatch());
obj.setProperty("text", "6");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/txtVEDA-VLAUFZ").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\03/ssubSUBSCREEN_BODY:SAPLV45W:4201/txtVEDA-VLAUFZ").toDispatch());
obj.setProperty("caretPosition", 1);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

// tab billing plan
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP/tabpT\\04").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\05").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\05").toDispatch());
obj.invoke("select");

// tab condition
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\06").toDispatch());
obj.invoke("select");

// isi table tab condition
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 4);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/lbl[1,9]").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/lbl[1,9]").toDispatch());
obj.setProperty("caretPosition", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 2);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/txtKOMV-KBETR").toDispatch());
obj.setProperty("text", "10000000");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[3]").toDispatch());
obj.invoke("press");

// tab texts and post link information
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09").toDispatch());
obj.invoke("select");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").toDispatch());
arg = new Variant[2];
arg[0] = new Variant("ZLIN");
arg[1] = new Variant("Column1");
obj.invoke("selectItem", arg);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").toDispatch());
arg = new Variant[2];
arg[0] = new Variant("ZLIN");
arg[1] = new Variant("Column1");
obj.invoke("ensureVisibleHorizontalItem", arg);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").toDispatch());
obj.setProperty("text", "    BA Date: 14.11.2023\r    Link:PT. Arta Flash Sintesa Nusantara \r    Bandwidth:2Gbps\r    Link ID:020C38616L1\r    CID:12-00349    \r\r");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\09/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").toDispatch());
arg = new Variant[2];
arg[0] = new Variant(115);
arg[1] = new Variant(115);
obj.invoke("setSelectionIndexes", arg);

// Tab additonal data A and select material group 3
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\10").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\11").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\12").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\13/ssubSUBSCREEN_BODY:SAPMV45A:4459/cmbVBAP-MVGR3").toDispatch());
obj.setProperty("key", "RRP");

// Tab Additional Data B
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14").toDispatch());
obj.invoke("select");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZLINKN").toDispatch());
obj.setProperty("text", "020C38616L1");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZSF_TRX").toDispatch());
obj.setProperty("text", "oBEw8Ze891447.0");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZCIRID").toDispatch());
obj.setProperty("text", "12-00349");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZSF_TRX_ITEM").toDispatch());
obj.setProperty("text", "CI0747914");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZBANDWID").toDispatch());
obj.setProperty("text", "2");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/ctxtVBAP-ZZUOM_BDWTH").toDispatch());
obj.setProperty("text", "12");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZSF_TRX_ITEM").toDispatch());
obj.invoke("setFocus");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\14/ssubSUBSCREEN_BODY:SAPMV45A:4462/subKUNDEN-SUBSCREEN_8459:SAPMV45A:8459/txtVBAP-ZZSF_TRX_ITEM").toDispatch());
obj.setProperty("caretPosition", 9);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

//Save
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-VAROPTION1").toDispatch());
obj.invoke("press");

obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/sbar").toDispatch());

// Retrieve the text property
String text2 = obj.getProperty("text")

// Check if a match is found
String contractNumber = sapHelpers.getNumber(text2)

if(contractNumber) {
	println("Contract Number : " + contractNumber)
}


//Exit
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[12]").toDispatch());
obj.invoke("press");

// Goto  VA42
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
obj.setProperty("text", "VA42");
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

// Go to detail
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VBELN").toDispatch());
obj.setProperty("caretPosition", 7);
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
obj.invoke("sendVKey", 0);

// klik button display header list
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").toDispatch());
obj.invoke("press");


// Tab Status
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11").toDispatch());
obj.invoke("select");

// Klik Object Status
obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\\11/ssubSUBSCREEN_BODY:SAPMV45A:4305/btnBT_KSTC").toDispatch());
obj.invoke("press");

// Edit Status to be Billing relesed
	// ini gw coba buat manual dlu
//obj = new ActiveXComponent(session.invoke("findById", "/app/con[0]/ses[0]/wnd[0]/usr/tabsTABSTRIP_0300/tabpANWS/ssubSUBSCREEN:SAPLBSVA:0302/tblSAPLBSVATC_E").toDispatch());
// ============================================================================================
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[3]").toDispatch());
//obj.invoke("press");
//
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[11]").toDispatch());
//obj.invoke("press");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-VAROPTION1").toDispatch());
//obj.invoke("press");
//
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/btn[3]").toDispatch());
//obj.invoke("press");
//
//
//// Goto  VA42
//WebUI.delay(30)
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/tbar[0]/okcd").toDispatch());
//obj.setProperty("text", "VA42");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
//obj.invoke("sendVKey", 0);
//
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/ctxtVBAK-VBELN").toDispatch());
//obj.setProperty("caretPosition", 7);
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
//obj.invoke("sendVKey", 0);
//
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\07").toDispatch());
//obj.invoke("select");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\\07/ssubSUBSCREEN_BODY:SAPMV45A:4409/subSUBSCREEN_TC:SAPMV45A:4922/subSUBSCREEN_BUTTONS:SAPMV45A:4052/btnBT_PFPL").toDispatch());
//obj.invoke("press");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/tbar[0]/btn[0]").toDispatch());
//obj.invoke("press");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\\07").toDispatch());
//obj.invoke("select");
//
////Close Application
//obj = new ActiveXComponent(session.invoke("findById", "wnd[0]").toDispatch());
//obj.invoke("close");
//obj = new ActiveXComponent(session.invoke("findById", "wnd[1]/usr/btnSPOP-OPTION1").toDispatch());
//obj.invoke("press");







	