/*
 * package com.example.Excel.Reports;
 * 
 * public class EXCEL { COLUMNs = new String[]{"Zone", "Region", "Depot"};
 * Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new
 * ByteArrayOutputStream();) { Sheet sheet =
 * workbook.createSheet("UTCL ORDER EPOD REPORT");
 * 
 * List<SalesReportEPODSummaryVO> dealerDataList = (List) dealerDataList1; Font
 * headerFont = workbook.createFont(); headerFont.setBold(true);
 * headerFont.setColor(IndexedColors.BLUE.getIndex());
 * 
 * CellStyle headerCellStyle = workbook.createCellStyle();
 * headerCellStyle.setFont(headerFont);
 * 
 * // Row for Header Row headerRow = sheet.createRow(0);
 * 
 * // Header for (int col = 0; col < COLUMNs.length; col++) { Cell cell =
 * headerRow.createCell(col); cell.setCellValue(COLUMNs[col]);
 * cell.setCellStyle(headerCellStyle); } int rowIdx = 1; for
 * (SalesReportEPODSummaryVO reportObject : dealerDataList) { Row row =
 * sheet.createRow(rowIdx++); int cellCount=0;
 * row.createCell(cellCount++).setCellValue(reportObject.getZoneDesc());
 * row.createCell(cellCount++).setCellValue(reportObject.getRegionDesc());
 * 
 * if(zrdType.equals(AnalyticsConstants.FIELD_REGION_CN)) {
 * row.createCell(cellCount++).setCellValue(reportObject.getDepotDesc()); }
 * if(zrdType.equals(AnalyticsConstants.FIELD_DEPOT_CN)) {
 * row.createCell(cellCount++).setCellValue(reportObject.getDepotDesc());
 * row.createCell(cellCount++).setCellValue(reportObject.getDealerCode());
 * row.createCell(cellCount++).setCellValue(reportObject.getDealerName());
 * row.createCell(cellCount++).setCellValue(reportObject.getDealerMobileNo());
 * 
 * } row.createCell(cellCount++).setCellValue(reportObject.getMonth());
 * 
 * if((zrdType.equals(AnalyticsConstants.FIELD_DEPOT_CN)))
 * row.createCell(cellCount++).setCellValue(reportObject.getIsOuaInstalled());
 * else
 * row.createCell(cellCount++).setCellValue(reportObject.getOuaInstalled());
 * 
 * row.createCell(cellCount++).setCellValue(reportObject.getTotalOrderBooked());
 * row.createCell(cellCount++).setCellValue(reportObject.getOrdersBookedOnOua())
 * ;
 * row.createCell(cellCount++).setCellValue(reportObject.getPercentDigitalOrder(
 * ));
 * row.createCell(cellCount++).setCellValue(reportObject.getRetailerOrders());
 * row.createCell(cellCount++).setCellValue(reportObject.
 * getTotalInvoiceGenerated());
 * row.createCell(cellCount++).setCellValue(reportObject.getTotalPrimaryInvoice(
 * )); row.createCell(cellCount++).setCellValue(reportObject.
 * getTotalSecondaryInvoice());
 * row.createCell(cellCount++).setCellValue(reportObject.
 * getEpodConfirmPrimarySource());
 * row.createCell(cellCount++).setCellValue(reportObject.
 * getEpodConfirmSecondarySource());
 * row.createCell(cellCount++).setCellValue(reportObject.getPercentEpodPrimary()
 * ); row.createCell(cellCount++).setCellValue(reportObject.
 * getTotalExPrimaryInvoice());
 * row.createCell(cellCount++).setCellValue(reportObject.
 * getTotalExSecondaryInvoice());
 * row.createCell(cellCount++).setCellValue(reportObject.
 * getEpodConfirmPrimarySourceEX());
 * row.createCell(cellCount++).setCellValue(reportObject.
 * getEpodConfirmSecondarySourceEX());
 * 
 * } workbook.write(out); logger.
 * info(" DownloadToExcel{} excelForSalesOrderEPODReport() method execution End"
 * ); return out.toByteArray(); }
 * 
 * }
 */