package com.xlaxiata.helper

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
import com.kms.katalon.core.configuration.RunConfiguration
import com.kms.katalon.core.context.TestCaseContext
import internal.GlobalVariable
import java.awt.Robot
import java.awt.Toolkit
import java.awt.Rectangle
import java.awt.image.BufferedImage
import javax.imageio.ImageIO
import java.text.SimpleDateFormat
import java.util.Date

public class SapHelpers {
	def static String getNumber (String text) {
		String number
		def match = (text =~ /(\d+)/)

		// Check if a match is found
		if (match.find()) {
			number = match.group(1)
		}

		return number
	}
	
	def sampleBeforeTestCase(TestCaseContext testCaseContext) {
		String testCaseId = testCaseContext.getTestCaseId()
	}

	def static takeScreenshootWindows (String spesificFolder) {
		// Get the current test case name
		def currentTestCaseName = RunConfiguration.getExecutionSource().toString().substring(RunConfiguration.getExecutionSource().toString().lastIndexOf("\\")+1)
		
		String reportsFolder = RunConfiguration.getProjectDir() + "/Reports/" + "Screenshot/${currentTestCaseName}/${spesificFolder}"
		
		// Create the folder if it doesn't exist
		new File(reportsFolder).mkdirs()
		
		def timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date())

		// Take a screenshot
		String screenshotLocation = reportsFolder + "/screenshot_${timestamp}.png"
		Robot robot = new Robot()
		Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize())
		BufferedImage screenFullImage = robot.createScreenCapture(screenRect)
		ImageIO.write(screenFullImage, "png", new File(screenshotLocation))
	}
}
