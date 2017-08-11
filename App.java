package com.excel.ExcelReadWrite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

public class App

{

	private static final String FILE_NAME = "/home/nikhal/Videos/Addition.xls";
	private static final String FILE_NAME2 = "/home/nikhal/Videos/Addition2.xls";

	static FormulaEvaluator evaluator;

	public static void main(String[] args) throws IOException {
		try {
			FileInputStream file = new FileInputStream(new File(FILE_NAME));

			HSSFWorkbook workbook = new HSSFWorkbook(file);
			HSSFSheet sheet = workbook.getSheetAt(0);
			// Cell cell = null;
			Row row = sheet.getRow(0);
			evaluator = workbook.getCreationHelper().createFormulaEvaluator();

			HSSFCell cell = (HSSFCell) row.createCell(1);

			cell.setCellValue(20);

			cell = (HSSFCell) row.getCell(2);
			String val = getCellValue(cell);
			System.out.println(val);

			file.close();

			FileOutputStream outFile = new FileOutputStream(new File(FILE_NAME2));
			workbook.write(outFile);
			outFile.close();

			

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	

	static String getCellValue(HSSFCell cell) {
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_NUMERIC:

			// if(cell.getColumnIndex()==18)
			// log.info("T>>>"+cell.getNumericCellValue());
			return String.valueOf(cell.getNumericCellValue());

		case HSSFCell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		case HSSFCell.CELL_TYPE_FORMULA:
			// if(cell.getColumnIndex()==18)
			// log.info("T>>> is formula"+cell.getNumericCellValue());
			evaluator.evaluateInCell(cell);
			// System.out.println(" Formula value: " + getCellValue(cell));
			return getCellValue(cell);
		case HSSFCell.CELL_TYPE_BLANK:
			return null;// return "Blank";
		case HSSFCell.CELL_TYPE_BOOLEAN:
			return null;
		default:
			return null;
		}
	}
}
