package com.example.Excel.Reports.Repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.example.Excel.Reports.Models.ReportExcel;

@Repository
public interface ExcelRepository extends JpaRepository<ReportExcel, String>{
	

}
