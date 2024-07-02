package bouguern.tuto.demo.service;

import java.util.List;

import bouguern.tuto.demo.entity.Course;
import bouguern.tuto.demo.entity.Employee;
import jakarta.servlet.http.HttpServletResponse;

public interface ReportService {
	
	public void generateExcel(HttpServletResponse response) throws Exception;

	public Employee saveEmployee(Employee employee);
	
	public List<Course> readExcelFile();

}
