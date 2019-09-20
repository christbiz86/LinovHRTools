package com.linov.excelengine.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.linov.excelengine.service.ExcelService;

@CrossOrigin(origins = "*")
@Controller
@RestController
@RequestMapping({"/xls"})
public class ExcelController {
	@Autowired
	ExcelService excelService;
	
	@GetMapping(value = "/write")
	public ResponseEntity<?> generateReportPdf() {
		try {
			excelService.writeXls();
			return ResponseEntity.status(HttpStatus.OK).body("Success");
		} catch (Exception e) {
			return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("F");
		}
	}
}
