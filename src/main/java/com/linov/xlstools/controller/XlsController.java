package com.linov.xlstools.controller;

import java.io.BufferedInputStream;
import java.io.InputStream;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.linov.xlstools.pojo.XlsReportPOJO;
import com.linov.xlstools.tools.XlsReader;
import com.linov.xlstools.tools.XlsWriter;

@CrossOrigin(origins = "*")
@Controller
@RestController
@RequestMapping({"/xls"})
public class XlsController {
	@Autowired
	XlsWriter xlsWriter;
	
	@Autowired
	XlsReader xlsReader;
	
	@PostMapping(value = "/write")
	public ResponseEntity<?> generateReportXls(@RequestBody XlsReportPOJO xlsReportPOJO) {
		try {
			xlsWriter.writeXls(xlsReportPOJO);
			return ResponseEntity.status(HttpStatus.OK).body("Success");
		} catch (Exception e) {
			return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("F");
		}
	}

	@PostMapping(value = "/read")
	public ResponseEntity<?> readReportXls(@RequestParam("file") MultipartFile file, 
			@RequestParam("start") String startCell, @RequestParam("end") String endCell,
			@RequestParam(value = "sheet", required = false) String sheetName) {
		try {
			InputStream inputStream =  new BufferedInputStream(file.getInputStream());
			return ResponseEntity.status(HttpStatus.OK).body(xlsReader.readXls(inputStream, sheetName, startCell, endCell));
		} catch (Exception e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("F");
		}
	}
}
