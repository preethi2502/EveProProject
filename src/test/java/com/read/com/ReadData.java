package com.read.com;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadData {

	public static void main(String[] args) throws Throwable {
		File f = new File("C:\\Users\\HARI\\eclipse-workspace\\DataDrivenConcepts\\DataFromProductOwner.xlsx");
		FileInputStream fis = new FileInputStream(f);
		Workbook wb = new XSSFWorkbook(fis);

		// to read data
		Sheet sheetAt = wb.getSheetAt(0);
		int row_size = sheetAt.getPhysicalNumberOfRows();

		// get datas using for loop
		for (int i = 0; i < row_size; i++) {
			Row row = sheetAt.getRow(i);

			int cell_Size = row.getPhysicalNumberOfCells();

			for (int j = 0; j < cell_Size; j++) {
				Cell cell = row.getCell(j);
				
				CellType cellType = cell.getCellType();
				
				if (cellType.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				}
				
				else if (cellType.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					int value = (int) numericCellValue;  //narrowCasting
					
					System.out.println(value);
					
				}
				
				
			}

		}

	}

}
