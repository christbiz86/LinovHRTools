package com.linov.xlstools.controller;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.linov.xlstools.pojo.XlsReportPOJO;
import com.linov.xlstools.tools.XlsReader;
import com.linov.xlstools.tools.model.ExcelFooter;
import com.linov.xlstools.tools.model.ExcelHeader;
import com.linov.xlstools.tools.model.ExcelReport;
import com.linov.xlstools.tools.model.ExcelSheet;
import com.linov.xlstools.tools.model.ExcelTable;

@CrossOrigin(origins = "*")
@Controller
@RestController
@RequestMapping({"/xls"})
public class XlsController {
	@Autowired
	XlsReader xlsReader;
	
	@GetMapping(value = "/write")
	public ResponseEntity<?> generateReportXls() {
		try {
			ExcelReport report = new ExcelReport("ReportTest");
			report.addSheet("Test Sheet 1");
			
			ExcelSheet sheet = report.getSheet(0);
			ExcelHeader header = sheet.getHeader();
			ExcelTable table = sheet.getTable();
			ExcelFooter footer = sheet.getFooter();
			
			header.addText("Header 1");
			
			table.addField("Field 1");
			List<Object> record = new ArrayList<Object>();
			record.add("Data 1");
			
			table.addRecord(record);
			footer.addText("Footer 1");
			
			report.createFile();
			
			return ResponseEntity.status(HttpStatus.OK).body("Success");
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("F");
		}
	}

	@PostMapping(value = "/read")
	public ResponseEntity<?> readReportXls(@RequestParam("file") MultipartFile file, 
			@RequestParam(value = "start", required = false) String startCell, @RequestParam(value = "end", required = false) String endCell,
			@RequestParam(value = "sheet", required = false) String sheetName) {
		try {
			InputStream inputStream =  new BufferedInputStream(file.getInputStream());
			
			return ResponseEntity.status(HttpStatus.OK).body(xlsReader.readXls(inputStream));
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("F");
		}
	}
}
