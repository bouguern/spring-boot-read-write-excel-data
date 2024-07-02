package bouguern.tuto.demo.controller;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import bouguern.tuto.demo.entity.Course;
import bouguern.tuto.demo.entity.Employee;
import bouguern.tuto.demo.service.ReportService;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;

import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.ResponseBody;

@RequestMapping("/excel")
@RestController
@RequiredArgsConstructor
public class EmployeeController {

	private static final Logger logger = LoggerFactory.getLogger(EmployeeController.class);

	private final ReportService reportService;

	@PostMapping("employees")
	public ResponseEntity<Employee> createEmployee(@RequestBody Employee employee) {
		Employee savedEmployee = reportService.saveEmployee(employee);
		logger.info("Create an employee");
		return ResponseEntity.ok(savedEmployee);
	}

	@GetMapping("/write")
	public void generateExcelReport(HttpServletResponse response) throws Exception {

		response.setContentType("application/octet-stream");

		// Excel file will be generated and saved to C:\Users\admin\Downloads as
		// 'employee.xls'

		String headerKey = "Content-Disposition";
		String headerValue = "attachment;filename=employees.xls";

		response.setHeader(headerKey, headerValue);

		logger.info("Generate an Excel file");
		reportService.generateExcel(response);

		response.flushBuffer();
	}

	@GetMapping("/read")
	public @ResponseBody List<Course> readCSV() {
		logger.info("Read courses in Excel file and return courses in JSON format");
		return reportService.readExcelFile();
	}
}
