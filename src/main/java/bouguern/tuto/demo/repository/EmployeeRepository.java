package bouguern.tuto.demo.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import bouguern.tuto.demo.entity.Employee;

@Repository
public interface EmployeeRepository extends JpaRepository<Employee, Long> {

}
