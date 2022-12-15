package com.mycompany.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingToExcel {

	private String directory = "upload";
	
	public void writingToExcel() {
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Employee");
		Row header = sheet.createRow(0);
		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("EMP_ID");
		headerCell = header.createCell(1);
		headerCell.setCellValue("EMP_NAME");
		headerCell = header.createCell(2);
		headerCell.setCellValue("EMP_NO");
		Row content = sheet.createRow(1);
		Cell contentCell = content.createCell(0);
		contentCell.setCellValue("221");
		contentCell = content.createCell(1);
		contentCell.setCellValue("À¯ÇÏÁø");
		contentCell = content.createCell(2);
		contentCell.setCellValue("800808-1123341");
		
		try {
			Path filePath = Paths.get(directory, "temp.xlsx");
			Files.createFile(filePath);
			FileOutputStream output = new FileOutputStream(filePath.toString());
			workbook.write(output);
			workbook.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}
