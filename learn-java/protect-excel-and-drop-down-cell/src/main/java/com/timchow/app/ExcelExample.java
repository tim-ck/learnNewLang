/*
 * author: Tim Chow
 * Github: tim-ck
 */
package com.timchow.app;

import org.apache.poi.hssf.usermodel.*;
// import org.apache.poi.ss.usermodel.HSSFDataValidation;
import org.apache.poi.ss.util.CellRangeAddressList;

import org.apache.poi.ss.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelExample {

	public static void main(String[] args) {
		// Example of how it works
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Sheet1");
		// create a row with header "Name", "Age", "preference"
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Name");
		header.createCell(1).setCellValue("Age");
		header.createCell(2).setCellValue("Preference");
		// create a row with data "John", 20, "Blue"
		Row dataRow = sheet.createRow(1);
		dataRow.createCell(0).setCellValue("John");
		dataRow.createCell(1).setCellValue(20);
		dataRow.createCell(2).setCellValue("Blue");

		setCellRangeAddressToDropDownCell(workbook, "Sheet1", 1, 2, 2, 2, new String[] { "Red", "Blue", "Green" });
		protectSheet(workbook, "Sheet1", new int[] { 2 }, "password");

		// save excel to project directory
		try {
			FileOutputStream out = new FileOutputStream("workbook.xlsx");
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	/**
	 * protectSheet but unlock list specified col
	 * The sheetName argument must match the sheet name in the workbook argument.
	 * row and col argument define number of column/row freeze in the sheet. set to
	 * 0 if no row/ col needed to be freeze.
	 * 
	 *
	 * @param wb              org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName       name of the sheet
	 * @param listOfUnlockCol list of Column index that need to be editable for
	 *                        user.
	 * @param password        password protection for the locked sheet
	 */
	public static void protectSheet(Workbook wb, String sheetName, int[] listOfUnlockCol, String password) {
		Sheet sheet = wb.getSheet(sheetName);
		CellStyle unlockedCellStyle = wb.createCellStyle();
		unlockedCellStyle.setLocked(false);
		// apply style to all cell in col
		for (int col : listOfUnlockCol) {
			for (int i = 0; i <= sheet.getLastRowNum(); i++) {
				sheet.getRow(i).getCell(col).setCellStyle(unlockedCellStyle);
			}
		}
		sheet.protectSheet(password);
	}

	/**
	 * Freeze panel in the sheet.
	 * The sheetName argument must match the sheet name in the workbook argument.
	 * row and col argument define number of column/row freeze in the sheet. set to
	 * 0 if no row/ col needed to be freeze.
	 * 
	 *
	 * @param wb        org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName name of the sheet
	 * @param colSplit  Horizontal position of split.
	 * @param rowSplit  Vertical position of split.
	 */
	public static void freezePanel(Workbook wb, String sheetName, int colSplit, int rowSplit) {
		Sheet sheet = wb.getSheet(sheetName);
		sheet.createFreezePane(colSplit, rowSplit);
		return;
	}

	/**
	 * Creates drop down list cell in sheet with defined position. Last row is set
	 * to last row index in the sheet.
	 * <p>
	 * if your sheet contain header, the first row should set to 1
	 *
	 * @param wb                 org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName          name of the sheet
	 * @param firstRow           Start Vertical position of split.
	 * @param firstCol           Start Horizontal position of split.
	 * @param firstCol           End Horizontal position of split.
	 * @param dataValidationList String array of drop down list data
	 */
	public static void setCellRangeAddressToDropDownCell(Workbook wb, String sheetName,
			int firstRow, int firstCol, int lastCol, String[] dataValidationList) {
		Sheet sheet = wb.getSheet(sheetName);
		sheet.getLastRowNum();
		setCellRangeAddressToDropDownCell(wb, sheetName, firstRow, sheet.getLastRowNum(), firstCol, lastCol,
				dataValidationList);
	}

	/**
	 * Creates drop down list cell in sheet with defined position
	 *
	 * <p>
	 * if your sheet contain header, the first row should set to 1
	 * 
	 * @param wb                 org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName          name of the sheet
	 * @param firstRow           Start Vertical position of split.
	 * @param lastRow            End Vertical position of split.
	 * @param firstCol           Start Horizontal position of split.
	 * @param firstCol           End Horizontal position of split.
	 * @param dataValidationList String array of drop down list data
	 */
	public static void setCellRangeAddressToDropDownCell(Workbook wb, String sheetName,
			int firstRow, int lastRow, int firstCol, int lastCol, String[] dataValidationList) {
		Sheet sheet = wb.getSheet(sheetName);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		DVConstraint constraint = DVConstraint.createExplicitListConstraint(dataValidationList);
		HSSFDataValidation validation = new HSSFDataValidation(addressList, constraint);
		sheet.addValidationData(validation);
	}
}