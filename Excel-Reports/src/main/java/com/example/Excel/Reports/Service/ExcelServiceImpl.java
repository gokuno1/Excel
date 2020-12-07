package com.example.Excel.Reports.Service;

import java.awt.Desktop;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.example.Excel.Reports.Models.NReport;
import com.example.Excel.Reports.Models.ReportExcel;
import com.example.Excel.Reports.Repository.ExcelRepository;
import com.example.Excel.Reports.Repository.NExcexlRepository;
import com.google.common.collect.Lists;

@Service
public class ExcelServiceImpl implements ExcelService{
	
	@Autowired
	private ExcelRepository excelRepository;
	
	private NExcexlRepository nexcexlRepository; 

	@Override //List<XSSFWorkbook>
	public List<XSSFWorkbook> generateExcelReport(List<ReportExcel> reportData) throws IOException {
		// TODO Auto-generated method stub
		List<XSSFWorkbook> workbooks = new ArrayList<XSSFWorkbook>();
		String[] COLUMNs=new String[]{"ZONE","REGION","DEPOT","ROLE","SAP_CODE","MOBILE_NUMBER","COMPANY_NAME","FIRST_NAME",
				"LAST_NAME","LAST_LOGIN_DATE","FIRST_TIME_LOGIN_DATE"};
		
		try (XSSFWorkbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream();) {
		Sheet sheet = workbook.createSheet("CB_SOUTH_A");
		
		List<ReportExcel> dealerDataList1 = reportData;
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
		//for(List<ReportExcel> dealerDataList: getDataInChunk(dealerDataList1) ) {
		for (ReportExcel reportObject : dealerDataList1) {
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
		out.close();
		return workbooks;
		//return out.toByteArray();
	}
		
	}

	@Override
	public List<ReportExcel> getAllData() {
		// TODO Auto-generated method stub
		List<ReportExcel> reportData = new ArrayList<>();
		reportData = excelRepository.findAll();
		return reportData;
	}

	public void writeWorkBooks(List<XSSFWorkbook> wbs) {
		String fileName = "SOUTH_DATA";
        FileOutputStream out;
        try {
            for (int i = 0; i < wbs.size(); i++) {
                String newFileName = fileName.substring(0, fileName.length() - 5);
                out = new FileOutputStream(new File(newFileName + "_" + (i + 1) + ".xlsx"));
                wbs.get(i).write(out);
                out.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
	
	public static  List<List<ReportExcel>> getDataInChunk(List<ReportExcel> inputList)
	{
	
		List<List<ReportExcel>> chunkDataList = Lists.partition(inputList, 1000);
		return chunkDataList;
	}

	@Override
	public NReport readExcelReport(MultipartFile file) throws IOException {
		// TODO Auto-generated method stub
		
		NReport report = new NReport();
		InputStream newFile = file.getInputStream();
		XSSFWorkbook workbook = new XSSFWorkbook(newFile);
		XSSFSheet sheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = sheet.iterator();

		DataFormatter formatter = new DataFormatter();
		
		
		while(rowIterator.hasNext())
		{
			Row row = rowIterator.next();
				report.setZone(row.getCell(0).getStringCellValue());
				report.setRegion(row.getCell(1).getStringCellValue());
				report.setDepot(row.getCell(2).getStringCellValue());
				report.setRole(row.getCell(3).getStringCellValue());
				report.setSapcode(row.getCell(4).getStringCellValue());
				report.setMobileNo("number");
				report.setCompanyName(row.getCell(6).getStringCellValue());
				report.setFirstName(row.getCell(7).getStringCellValue());
				report.setLastName(row.getCell(8).getStringCellValue());
				report.setFirstTimeLoginDate(String.valueOf(row.getCell(9).getDateCellValue()));
				report.setLastLoginDate(String.valueOf(row.getCell(10).getDateCellValue()));
			
		}
		
		return report;
	}
	
	private List<SXSSFWorkbook> splitWorkbook(XSSFWorkbook workbook) {

        List<SXSSFWorkbook> workbooks = new ArrayList<SXSSFWorkbook>();

        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = (SXSSFSheet) wb.createSheet();

        Row newRow;
        SXSSFCell newCell;

        int rowCount = 0;
        int colCount = 0;

        XSSFSheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            newRow = sh.createRow(rowCount++);

            /* Time to create a new workbook? */
            if (rowCount == 1000) {
                workbooks.add(wb);
                wb = new SXSSFWorkbook();
                sh = (SXSSFSheet) wb.createSheet();
                rowCount = 0;
            }

            for (Cell cell : row) {
                newCell = (SXSSFCell) newRow.createCell(colCount++);
                newCell = setValue(newCell, cell);

                CellStyle newStyle = wb.createCellStyle();
                newStyle.cloneStyleFrom(cell.getCellStyle());
                newCell.setCellStyle(newStyle);
            }
            colCount = 0;
        }

        /* Only add the last workbook if it has content */
        if (wb.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
            workbooks.add(wb);
        }
        return workbooks;
    }

	
	
	
}
