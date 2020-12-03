package com.example.Excel.Reports.Service;

import java.io.IOException;
import java.util.List;

import com.example.Excel.Reports.Models.ReportExcel;

public interface ExcelService {
	
	public List<ReportExcel> getAllData();
	public byte[] generateExcelReport(List<ReportExcel> reportData) throws IOException;

}
