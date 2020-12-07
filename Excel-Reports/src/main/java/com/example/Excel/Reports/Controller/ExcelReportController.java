package com.example.Excel.Reports.Controller;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.Excel.Reports.Models.NReport;
import com.example.Excel.Reports.Models.ReportExcel;
import com.example.Excel.Reports.Service.ExcelServiceImpl;

@RestController
public class ExcelReportController {
	
	@Autowired
	private ExcelServiceImpl excelService;
	
	@GetMapping(value = "/downloadExcel")
	public ResponseEntity<byte[]> generateExcel() throws IOException
	{
		List<ReportExcel> reportData = excelService.getAllData();
		byte [] excel = excelService.generateExcelReport(reportData);
		 return ResponseEntity.ok()
			        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + "CB_SOUTH"+".xlsx")
			        .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
			        .body(excel);
	}
	
	@PostMapping(value = "/uploadData")	
	public NReport uploadData(@RequestPart MultipartFile file) throws IOException
	{
		String fileName = file.getOriginalFilename();
		NReport report =  excelService.readExcelReport(file);
		return report;
	}

}
