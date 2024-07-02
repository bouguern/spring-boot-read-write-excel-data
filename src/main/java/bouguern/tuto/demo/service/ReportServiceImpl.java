package bouguern.tuto.demo.service;

import java.io.File;
import java.io.IOException;
import java.sql.Date;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import bouguern.tuto.demo.entity.Course;
import bouguern.tuto.demo.entity.Employee;
import bouguern.tuto.demo.repository.EmployeeRepository;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import lombok.RequiredArgsConstructor;

@Service
@RequiredArgsConstructor
public class ReportServiceImpl implements ReportService {

	private static final Logger logger = LoggerFactory.getLogger(ReportServiceImpl.class);

	private final EmployeeRepository employeeRepository;

	/**
	 * You must create "demo.xlsx" before testing this API
	 * And enter some data inside the file with 4 columns
	 */
	private static final String CSV_FILE_LOCATION = "C:/Users/admin/Documents/demo.xlsx";

	@Override
	public Employee saveEmployee(Employee employee) {
		return employeeRepository.save(employee);
	}

	@Override
	public void generateExcel(HttpServletResponse response) throws Exception {

		List<Employee> employees = employeeRepository.findAll();

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet("Employees Info");
		HSSFRow row = sheet.createRow(0);

		row.createCell(0).setCellValue("ID employee");
		row.createCell(1).setCellValue("First Name");
		row.createCell(2).setCellValue("Last Name");
		row.createCell(3).setCellValue("Started Date");

		// Create date cell style
		HSSFCellStyle dateCellStyle = workbook.createCellStyle();
		HSSFDataFormat dateFormat = workbook.createDataFormat();
		// dateCellStyle.setDataFormat(dateFormat.getFormat("yyyy-mm-dd"));
		dateCellStyle.setDataFormat(dateFormat.getFormat("dd-mm-yyyy"));

		int dataRowIndex = 1;

		for (Employee employee : employees) {
			HSSFRow dataRow = sheet.createRow(dataRowIndex);
			dataRow.createCell(0).setCellValue(employee.getEmployeeId());
			dataRow.createCell(1).setCellValue(employee.getFirstName());
			dataRow.createCell(2).setCellValue(employee.getLastName());

			// Set the date value with the date format style
			if(employee.getStartedDateInCompany() != null) {
				HSSFCell dateCell = dataRow.createCell(3);
				dateCell.setCellValue(Date.valueOf(employee.getStartedDateInCompany()));
				dateCell.setCellStyle(dateCellStyle);
			}

			dataRowIndex++;
		}

		// Auto-size columns
		for (int i = 0; i < 4; i++) {
			sheet.autoSizeColumn(i);
		}

		ServletOutputStream ops = response.getOutputStream();
		workbook.write(ops);
		workbook.close();
		ops.close();

	}

	@Override
	public List<Course> readExcelFile() {

		List<Course> courses = new ArrayList<>();

		Workbook workbook = null;
		try {
			// Creating a Workbook from an Excel file (.xls or .xlsx)
			workbook = WorkbookFactory.create(new File(CSV_FILE_LOCATION));

			// Retrieving the number of sheets in the Workbook
			logger.info("Number of sheets: " + workbook.getNumberOfSheets());

			// Print all sheets name
			workbook.forEach(sheet -> {
				logger.info("Title of sheet => " + sheet.getSheetName());

				// Create a DataFormatter to format and get each cell's value as String
				DataFormatter dataFormatter = new DataFormatter();

				// Loop through all rows and columns and create Course object
				int index = 0;
				for (Row row : sheet) {
					if (index++ == 0)
						continue;
					Course course = new Course();

					if (row.getCell(0) != null && row.getCell(0).getCellType() == CellType.NUMERIC) {
						course.setId((int) row.getCell(0).getNumericCellValue());
					}

					if (row.getCell(1) != null) {
						course.setName(dataFormatter.formatCellValue(row.getCell(1)));
					}

					Cell dateCell = row.getCell(2);
					if (DateUtil.isCellDateFormatted(dateCell)) {
						LocalDate date = dateCell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault())
								.toLocalDate();
						course.setDate(date);
					}

					if (row.getCell(3) != null && row.getCell(3).getCellType() == CellType.NUMERIC) {
						course.setNumber((int) row.getCell(3).getNumericCellValue());
					}
					courses.add(course);
				}
			});
		} catch (EncryptedDocumentException | IOException e) {
			logger.error(e.getMessage(), e);
		} finally {
			try {
				if (workbook != null)
					workbook.close();
			} catch (IOException e) {
				logger.error(e.getMessage(), e);
			}
		}

		return courses;
	}

}
