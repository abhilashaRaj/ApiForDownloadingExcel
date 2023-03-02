package com.createExceltest.testing;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;


@RestController
public class ExcelController {

	
	@Autowired
	private ExcelService excelservice;
	
	@PostMapping(value = "/getexcelformat")
	public String getExcel(@RequestBody List<Excel> excel){
		
		
		return excelservice.getExcelFormat(excel);
		
	}
}
