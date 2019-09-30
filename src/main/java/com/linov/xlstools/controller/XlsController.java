package com.linov.xlstools.controller;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.linov.xlstools.tools.XlsReader;
import com.linov.xlstools.tools.model.ExcelFooter;
import com.linov.xlstools.tools.model.ExcelHeader;
import com.linov.xlstools.tools.model.ExcelReport;
import com.linov.xlstools.tools.model.ExcelSheet;
import com.linov.xlstools.tools.model.ExcelStyle;
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
			ExcelReport report = new ExcelReport("Test");
			report.addSheet("Test Sheet 1");
			
			ExcelSheet sheet = report.getSheet(0);
			ExcelHeader header = sheet.getHeader();
			ExcelTable table = sheet.getTable();
			ExcelFooter footer = sheet.getFooter();
			
			header.addText("Header 1");
			header.addText("Header 2");
			header.addText("Header 3");
			header.addText("Header 4");
			header.addText(null);
			header.removeText(header.getText(2));
			
			table.addField("Field 1");
			table.addField("Field 2");
			table.addField("Field 3");
			table.addField("Field 4");
			
			List<Object> record1 = new ArrayList<Object>();
			record1.add("Data 1");
			record1.add(1);
			record1.add(LocalDate.now());
			record1.add(true);
			
			List<Object> record2 = new ArrayList<Object>();
			record2.add("Data 2");
			record2.add(2);
			record2.add(LocalDate.now());
			record2.add(false);
			
			table.addRecord(record1);
			table.addRecord(record2);

			ExcelStyle dateStyle1 = table.getRecord(0).getText("Field 3").getStyle();
			dateStyle1.setDataFormat("dd/MM/yyyy");
			ExcelStyle dateStyle2 = table.getRecord(1).getText("Field 3").getStyle();
			dateStyle2.setDataFormat("dd/MM/yyyy");
			
			footer.addText("");
			footer.addText("Footer 1");
			footer.addText("Footer 2");
			footer.addText("Footer 3");
			footer.addText("Footer 4");
			footer.removeText(footer.getText(3));
			
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
