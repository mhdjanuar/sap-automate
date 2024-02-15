package com.xlaxiata.pageobjects.web

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
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import org.openqa.selenium.Keys as Keys

import internal.GlobalVariable

public class CreateMaterial {
	def login () {
		WebUI.setText(findTestObject('Web/InputUsername'), GlobalVariable.username)
		WebUI.setText(findTestObject('Web/InputPassword'), GlobalVariable.password)
		WebUI.click(findTestObject('Web/BtnLogin'))
	}

	def inputPersonalNo () {
		WebUI.setText(findTestObject('Web/InputPersonalNo'), "90007061")
		WebUI.sendKeys(findTestObject('Web/InputPersonalNo'), Keys.chord(Keys.ENTER))
		
	}

	def inputRequestNote (String reqNotes) {
		WebUI.setText(findTestObject('Web/TextAreaRequestNote'), "Create Material with type - ${reqNotes}")
		WebUI.click(findTestObject('Web/BtnNew'))
	}
	
	def inputMaterialNo () {
		WebUI.setText(findTestObject('Web/InputMaterialNo'), "SP10GB-JKT")
		WebUI.sendKeys(findTestObject('Web/InputMaterialNo'), Keys.chord(Keys.ENTER))
	}

	def selectWithArrowDown (TestObject target, Integer countArrowDown) {
		WebUI.click(target)
		
		for (int i = 1; i <= countArrowDown; i ++) {
			WebUI.sendKeys(target, Keys.chord(Keys.ARROW_DOWN))
		}
	
		WebUI.sendKeys(target, Keys.chord(Keys.ENTER))
	}

	def select (TestObject target, TestObject value) {
		WebUI.click(target)
		WebUI.click(value)
	}

	def tabClasificationStep (TestObject valueSelectNoun, Integer countModifier) {
		// Select Noun
		select(findTestObject('Web/SelectNoun'), valueSelectNoun)
			
		// Input Modifier
		selectWithArrowDown(findTestObject('Web/SelectModifier'), countModifier)
	}

	def tabBasicStep (
		TestObject valueMaterialGroup, 
		TestObject valueDivision, 
		String valueExternalMaterial, 
		TestObject valueLaboratory = null
	) {
		// Goto Tab Basic
		WebUI.click(findTestObject('Web/TabBasic'))
		
		//Input Material Group
		select(findTestObject('Web/SelectMaterialGroup'), valueMaterialGroup)
		
		//Input Division
		select(findTestObject('Web/SelectDivision'), valueDivision)
		
		//Input External Material Group
		WebUI.setText(findTestObject('Web/InputExternalMaterialGroup'), valueExternalMaterial)
		WebUI.sendKeys(findTestObject('Web/InputExternalMaterialGroup'), Keys.chord(Keys.ENTER))
		
		if(valueLaboratory) {
			//Input Laboratory
			select(findTestObject('Web/SelectLaboratory'), findTestObject('Web/ChooseLaboratory'))
		}
	}

	def tabAccountingStep (String price, String valuationClass) {
		// Goto Tab Accounting
		WebUI.click(findTestObject('Web/TabAccounting'))
		WebUI.setText(findTestObject('Web/InputPrice'), price)
		WebUI.setText(findTestObject('Web/InputValuationClass'), valuationClass)
	}

	def tabSalesAndPlant(
		String valuePlan,
		String valueSalesOrganization,
		TestObject valueGeneralItem,
		TestObject valueItem,
		TestObject valueSelectSalesOrganization,
		TestObject valueDistribution,
		TestObject valueAccountAssignment,
		TestObject valueMaterialGroupPacking = null,
		TestObject valueSerialNumber = null,
		TestObject targetShareLocation
	) {
		// Goto Tab Sales & Plant
		WebUI.click(findTestObject('Web/TabSalesPlant'))
		
		// Input field sales & plant
		WebUI.setText(findTestObject('Web/InputPlant'), valuePlan)
		WebUI.sendKeys(findTestObject('Web/InputPlant'), Keys.chord(Keys.ENTER))
		
		WebUI.setText(targetShareLocation, valueSalesOrganization)
		WebUI.sendKeys(targetShareLocation, Keys.chord(Keys.ENTER))
		
		//Select General Item
		select(findTestObject('Web/SelectGeneralItem'), valueGeneralItem)
		
		//Select Item Category
		select(findTestObject('Web/SelectItem'), valueItem)
		
		//Select Sales Organization
		select(findTestObject('Web/SelectSalesOrganization'), valueSelectSalesOrganization)
		
		//Select Distribution Channel
		select(findTestObject('Web/SelectDistributionChannel'), valueDistribution)
		
		//Select Account Assignment
		select(findTestObject('Web/SelectAccountAssignment'), valueAccountAssignment)
		
		if(valueMaterialGroupPacking) {
			//Select Material Group Packing
			select(findTestObject('Web/SelectMaterialGroupPacking'), valueMaterialGroupPacking)
		}
		
		if(valueSerialNumber) {
			//Select Serial Number
			select(findTestObject('Web/SelectSerialNumber'), valueSerialNumber)
		}
	}

	def submit () {
		WebUI.click(findTestObject('Web/BtnSave'))
		
		boolean isButtonOk = WebUI.verifyElementPresent(findTestObject('Web/BtnOk'), 5, FailureHandling.CONTINUE_ON_FAILURE)
		
		if(isButtonOk) {
			WebUI.click(findTestObject('Web/BtnOk'))
//			WebUI.click(findTestObject('Web/BtnSaveDraft'))
//			WebUI.click(findTestObject('Web/BtnSubmit'))
		}
	}
	
	// ================================ Main Function
	
	def zeqpAssetXLEquipment () {
		tabClasificationStep(findTestObject('Web/ChooseOptionNoun'), 5)
		
		tabBasicStep(
			findTestObject('Web/ChooseMaterialGroupZ0001'),
			findTestObject('Web/ChooseDivision'),
			"NEW005",
			findTestObject('Web/ChooseLaboratory'))
		
		tabAccountingStep("1200000", "3201")
		
		tabSalesAndPlant(
			"0010",
			"0017",
			findTestObject('Web/ChooseGeneralItemCategory'),
			findTestObject('Web/ChooseItemCategory'),
			findTestObject('Web/ChooseSalesOrganization'),
			findTestObject('Web/ChooseDistributionChannel'),
			findTestObject('Web/ChooseAccountAssignment'),
			findTestObject('Web/ChooseMaterialGroupPacking'),
			findTestObject('Web/ChooseSerialNumber'),
			findTestObject('Web/InputSalesOrganization')
		)
		
		
		submit()
		
		//Back
		WebUI.click(findTestObject('Web/BtnBack'))
		WebUI.click(findTestObject('Web/IconBack'))
	}
	
	def zfgtNonAssetXL () {
		tabClasificationStep(findTestObject('Web/SelectOptionNounANT'), 7)
		
		tabBasicStep(
			findTestObject('Web/ChooseMaterialGroupZ0004'),
			findTestObject('Web/ChooseDivision'),
			"NEW005",
			findTestObject('Web/ChooseLaboratory'))
		
		tabAccountingStep("2300000", "3104")
		
		tabSalesAndPlant(
			"0010",
			"0011",
			findTestObject('Web/ChooseGeneralItemCategory'),
			findTestObject('Web/ChooseItemCategory'),
			findTestObject('Web/ChooseSalesOrganization'),
			findTestObject('Web/ChooseDistributionChannel'),
			findTestObject('Web/ChooseAccountAssignment'),
			null,
			null,
			findTestObject('Web/InputSalesOrganization'))
		
		submit()
		
		//Back
		WebUI.click(findTestObject('Web/BtnBack'))
		WebUI.click(findTestObject('Web/IconBack'))
	}
	
	def zhawXLTradingGoods () {
		tabClasificationStep(findTestObject('Web/SelectOptionNounSimcard'), 1)
		
		tabBasicStep(
			findTestObject('Web/ChooseMaterialGroupZ0001'),
			findTestObject('Web/ChooseDivision'),
			"NEW005",
			findTestObject('Web/ChooseLaboratory'))
		
		tabAccountingStep("2430000", "3104")
		
		tabSalesAndPlant(
			"0010",
			"0011",
			findTestObject('Web/ChooseGeneralItemCategory'),
			findTestObject('Web/ChooseItemCategory'),
			findTestObject('Web/ChooseSalesOrganization'),
			findTestObject('Web/ChooseDistributionChannel'),
			findTestObject('Web/ChooseAccountAssignment'),
			null,
			null,
			findTestObject('Web/InputShareLocation')
		 )
		 
		submit()
		
		//Back
		WebUI.click(findTestObject('Web/BtnBack'))
		WebUI.click(findTestObject('Web/IconBack'))
	}
	
	def znwsAssetXlNonEquip () {
		tabClasificationStep(findTestObject('Web/ChooseOptionNoun'), 2)
		
		tabBasicStep(
			findTestObject('Web/ChooseMaterialGroupZ0001'),
			findTestObject('Web/ChooseDivision'),
			"NEW005",
			findTestObject('Web/ChooseLaboratory'))
		
		tabAccountingStep("1300000", "3104")
		
		tabSalesAndPlant(
			"0010",
			"0020",
			findTestObject('Web/ChooseGeneralItemCategory'),
			findTestObject('Web/ChooseItemCategory'),
			findTestObject('Web/ChooseSalesOrganization'),
			findTestObject('Web/ChooseDistributionChannel'),
			findTestObject('Web/ChooseAccountAssignment'),
			findTestObject('Web/ChooseMaterialGroupPacking'),
			null,
			findTestObject('Web/InputSalesOrganization')
		 )
		
		submit()
		
		//Back
		WebUI.click(findTestObject('Web/BtnBack'))
		WebUI.click(findTestObject('Web/IconBack'))
	}
	
	def zserService () {
		tabClasificationStep(findTestObject('Web/SelectOptionNounSimcard'), 1)
		
		tabBasicStep(
			findTestObject('Web/ChooseMaterialGroupZ0004'),
			findTestObject('Web/ChooseDivision'),
			"NEW005",
			null
		)
		
		tabSalesAndPlant(
			"0010",
			"0020",
			findTestObject('Web/SelectOptionGeneralItemZSER'),
			findTestObject('Web/SelectOptionItemZSER'),
			findTestObject('Web/SelectOptionSalesOrganizationZSER'),
			findTestObject('Web/SelectOptionDistribusiChannelZSER'),
			findTestObject('Web/SelectOptionAccountAssigmentZSER'),
			null,
			null,
			findTestObject('Web/InputStorageLocationZSER')
		 )
		
		submit()
		
		//Back
		WebUI.click(findTestObject('Web/BtnBack'))
		WebUI.click(findTestObject('Web/IconBack'))
	}
	
	def zsfwXlSoftware () {
		tabClasificationStep(findTestObject('Web/SelectOptionNounFiller'), 2)
		
		tabBasicStep(
			findTestObject('Web/ChooseMaterialGroupZ0001'),
			findTestObject('Web/ChooseDivision'),
			"NEW003",
			null
		)
		
		tabAccountingStep("1300000", "3109")
		
		tabSalesAndPlant(
			"0010",
			"0020",
			findTestObject('Web/SelectOptionGeneralItemZSER'),
			findTestObject('Web/SelectOptionItemZSER'),
			findTestObject('Web/SelectOptionSalesOrganizationZSER'),
			findTestObject('Web/SelectOptionDistribusiChannelZSER'),
			findTestObject('Web/SelectOptionAccountAssigmentZSER'),
			null,
			null,
			findTestObject('Web/InputSalesOrganization')
		 )
		
		submit()
		
		//Back
		WebUI.click(findTestObject('Web/BtnBack'))
		WebUI.click(findTestObject('Web/IconBack'))
	}
 
}
