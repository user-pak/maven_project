package com.mycompany.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class InsertEmployeeIntoOracleFromExcel {

	public void insertData() {
		Workbook workbook = null;
		Connection conn = null;
		PreparedStatement pstmt = null;
		int result = 0;
		try {
			FileInputStream file =new FileInputStream("upload/통합 문서 3 (1).xlsx");
			Class.forName("oracle.jdbc.driver.OracleDriver");
			conn = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:XE", "EXCEL", "EXCEL");
			String sql = "INSERT INTO EMPLOYEE VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";
			pstmt = conn.prepareStatement(sql);
			workbook = new XSSFWorkbook(file);
			Sheet sheet = workbook.getSheetAt(0);

			for(int j = 1; j < sheet.getLastRowNum(); j++) {
				Row row = sheet.getRow(j);
				if(row == null) continue;
				for(int i = 0; i < row.getLastCellNum() ; i++) {
					if(row.getCell(i) == null) continue;
					switch(row.getCell(i).getCellType()) {
					case STRING :
						pstmt.setString(i+1, row.getCell(i).getStringCellValue()); break;
					case NUMERIC :
						if(DateUtil.isCellDateFormatted(row.getCell(i))) {
							pstmt.setDate(i+1, new Date(row.getCell(i).getDateCellValue().getTime())); break;
						}else {
							pstmt.setDouble(i+1, row.getCell(i).getNumericCellValue()); break;
						}
					default :
						pstmt.setString(i+1, null); break;
					}		
				}
				result += pstmt.executeUpdate();
			}

		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} finally {
			if(pstmt != null)
				try {
					pstmt.close();
				} catch (SQLException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if(conn != null)
				try {
					conn.close();
				} catch (SQLException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			if(workbook != null)
				try {
					workbook.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		}
		System.out.println(result);
	}
}
