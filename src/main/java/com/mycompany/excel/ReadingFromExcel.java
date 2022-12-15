package com.mycompany.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingFromExcel {
	
	public FileInputStream readingFromExcel() {
		
		FileInputStream file = null;
		try {
			file = new FileInputStream("upload/GROUP_BY 수행순서.xlsx");
			Workbook workbook = new XSSFWorkbook(file);
			Sheet sheet = workbook.getSheetAt(0);
			Map<Integer, List<String>> data = new HashMap<Integer, List<String>>();
			int i = 0;
			for(Row row : sheet) {
				data.put(i, new ArrayList<String>());
				for(Cell cell : row) {
					switch(cell.getCellType()) {
					case STRING :
						data.get(new Integer(i)).add(cell.getRichStringCellValue().toString()); break;
					case NUMERIC :
						if(DateUtil.isCellDateFormatted(cell)) {
							data.get(new Integer(i)).add(cell.getDateCellValue()+"");
						}else {
							data.get(new Integer(i)).add(cell.getNumericCellValue()+"");
						}
						break;
					case BOOLEAN :
						data.get(new Integer(i)).add(cell.getBooleanCellValue()+""); break;
					case FORMULA :
						data.get(new Integer(i)).add(cell.getCellFormula()+""); break;
					default:
						data.get(new Integer(i)).add("");
						break;
					}
				}
				i++;
			}
			
//			Iterator<Map.Entry<Integer, List<String>>> entries = data.entrySet().iterator();
//			while (entries.hasNext()) {
//			    Entry<Integer, List<String>> entry = entries.next();
//			    System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());
//			}
			for (List<String> value : data.values()) {
		    System.out.println("Value = " + value);
			}
//			for (Entry<Integer, List<String>> entry : data.entrySet()) {
//			    System.out.println("Key = " + entry.getKey() + ", Value = " + entry.getValue());
//			}	
			workbook.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}	

		return file;
	}	
}
