package com.example.Excel.Reports;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.*;

public class ReportExcel {

    @SuppressWarnings("null")
    public static void main(String[] args) throws IOException {
        
        int sc=1;
        Row r; Cell c;
        int i;
        int n=100000;//total result rows
        int rc=50000;// rows per book
        int cou=1,j=0,data=1;
        int doc=0;
        FileOutputStream outputStream = null;
        
        XSSFWorkbook wbo = new XSSFWorkbook();
        XSSFSheet s = null ;
        
        for(i=0;i<=n;i++)
        {
            if(i%rc == 0){
                s=wbo.createSheet();
            }
            if(i!=0 && i%rc == 0)
            {
                
                outputStream = new FileOutputStream("C:/Users/1239320/Desktop/Multi Sheet"+doc+ ".xlsx");
                wbo.write(outputStream);
                outputStream.flush();
                wbo.close();
                System.gc();
                doc++;
                cou=1;			
            }
            r = s.createRow(cou);
            for( j=0;j<15;j++)
            {
                c = r.createCell(j); 
                c.setCellValue(data);			
            }
            System.out.println(data);
            data++;
            cou++;
        }
        
        
        System.out.println("Completed");
    }
    
    
}