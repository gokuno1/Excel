package com.example.Excel.Reports.Service;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.web.servlet.MultipartAutoConfiguration;
import org.springframework.web.multipart.MultipartFile;

import com.example.Excel.Reports.Models.NReport;
import com.example.Excel.Reports.Models.ReportExcel;

public interface ExcelService {
	
	public List<ReportExcel> getAllData();
	public List<XSSFWorkbook> generateExcelReport(List<ReportExcel> reportData) throws IOException;
	public NReport readExcelReport(MultipartFile file) throws IOException;
//List<XSSFWorkbook>
}
