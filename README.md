# Reading and Writing Excel Data in Spring Boot with Apache POI

This project demonstrates how to read and write Excel data in a Spring Boot 3 application using Apache POI. It includes endpoints to create employee records, generate Excel reports, and read data from Excel files.


## Dependencies
Add the following dependencies to your `pom.xml` file:
```xml
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.0.0</version>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>5.0.0</version>
</dependency>
```
## Prerequisites

- JDK 17 or higher
- Spring Boot 3
- Git (optional, for cloning the repository)

## Getting Started

Follow these instructions to get a copy of the project up and running on your local machine for development and testing purposes.

### Clone the Repository

```bash
	git clone https://github.com/bouguern/spring-boot-read-write-excel-data.git
	
	cd spring-boot-read-write-excel-data
	
	mvn clean install

	mvn spring-boot:run
```
	
## API Endpoints

### 1. Create a new employee

- **Endpoint:** `POST http://localhost:9095/excel/employees`
- **Description:** This endpoint creates a new employee.
- **Request:**
  - **Content-Type:** application/json
  - **Body:**
    ```json
    {
        "firstName": "Mohamed",
    	"lastName": "Bouguern",
    	"startedDateInCompany": "2023-01-01"
    }
  - **Response:**
   - **Status Code:** 200 OK with the created employee object.
   
### 2. Generate Excel Report

- **Endpoint:** `GET http://localhost:9095/excel/write`
- **Description:** Generates an Excel file.
  - **Response:**
   - Generates an Excel file employees.xls and downloads it.
  
### 3. Read Data from Excel File

- **Endpoint:** `GET http://localhost:9095/excel/read`
- **Description:** Generates an Excel file.
  - **Response:**
   - 200 OK with a list of courses in JSON format.

## Feedback and Contributions

Feedback and contributions are welcome! Feel free to fork the repository, open issues, 
or submit pull requests to improve this example and make it even more useful for the community.
