package bouguern.tuto.demo.entity;

import java.time.LocalDate;

import lombok.Data;

@Data
public class Course {

	private int id;
	private String name;
	private LocalDate date;
	private int number;
}