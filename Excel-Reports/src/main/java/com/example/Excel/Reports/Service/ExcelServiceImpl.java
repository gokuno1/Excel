package com.example.Excel.Reports.Service;

import java.awt.Desktop;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.Excel.Reports.Models.ReportExcel;
import com.example.Excel.Reports.Repository.ExcelRepository;

@Service
public class ExcelServiceImpl implements ExcelService{
	
	@Autowired
	private ExcelRepository excelRepository;

	@Override
	public byte[] generateExcelReport(List<ReportExcel> reportData) throws IOException {
		// TODO Auto-generated method stub
		
		String[] COLUMNs=new String[]{"ZONE","REGION","DEPOT","ROLE","SAP_CODE","MOBILE_NUMBER","COMPANY_NAME","FIRST_NAME",
				"LAST_NAME","LAST_LOGIN_DATE","FIRST_TIME_LOGIN_DATE"};
		
		try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
		Sheet sheet = workbook.createSheet("CB_SOUTH_A");
		
		List<ReportExcel> dealerDataList = reportData;
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setColor(IndexedColors.BLUE.getIndex());

		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);

		// Row for Header
		Row headerRow = sheet.createRow(0);

		// Header
		for (int col = 0; col < COLUMNs.length; col++) {
			Cell cell = headerRow.createCell(col);
			cell.setCellValue(COLUMNs[col]);
			cell.setCellStyle(headerCellStyle);
		}
		int rowIdx = 1;
		for (ReportExcel reportObject : dealerDataList) {
			Row row = sheet.createRow(rowIdx++);
			int cellCount=0;
			row.createCell(cellCount++).setCellValue(reportObject.getZone());				
			row.createCell(cellCount++).setCellValue(reportObject.getRegion());
			row.createCell(cellCount++).setCellValue(reportObject.getDepot());
			row.createCell(cellCount++).setCellValue(reportObject.getRole());
			row.createCell(cellCount++).setCellValue(reportObject.getSapcode());
			row.createCell(cellCount++).setCellValue(reportObject.getMobileNo());
			row.createCell(cellCount++).setCellValue(reportObject.getCompanyName());
			row.createCell(cellCount++).setCellValue(reportObject.getFirstName());
			row.createCell(cellCount++).setCellValue(reportObject.getLastName());
			row.createCell(cellCount++).setCellValue(reportObject.getLastLoginDate());
			row.createCell(cellCount++).setCellValue(reportObject.getFirstTimeLoginDate());
		}
		workbook.write(out);
		return out.toByteArray();
	}
	}

	@Override
	public List<ReportExcel> getAllData() {
		// TODO Auto-generated method stub
		List<ReportExcel> reportData = new ArrayList<>();
		reportData = excelRepository.findAll();
		return reportData;
	}

}
