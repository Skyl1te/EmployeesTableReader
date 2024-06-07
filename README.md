# EmployeesTableReader

## Overview

**EmployeesTableReader** is a C# application designed to read and manipulate employee data stored in an Excel file. Utilizing the `EPPlus` library for Excel file operations, this project provides a range of functionalities to manage employee records efficiently. Key features include retrieving employees based on profession, calculating average salaries, identifying the highest-paid employee, and more.

## Features

- **Retrieve All Employees**: Fetches all employee records from the Excel file.
- **Get Employee by ID**: Finds and returns an employee record by their ID.
- **Get Employees by Profession**: Retrieves a list of employees filtered by a specified profession.
- **Get Unofficially Employed Employees**: Lists employees who are not officially employed.
- **Calculate Average Salary**: Computes the average salary of all employees.
- **Find Highest Paid Employee**: Identifies and returns the employee with the highest salary.
- **Additional Filtering**:
  - Employees with salary above a specified threshold.
  - Employees older or younger than a specified age.
  - Employees with specific diseases.
- **CRUD Operations**:
  - Add a new employee to the Excel file.
  - Update an existing employee's information.
  - Delete an employee record.

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/yourusername/EmployeesTableReader.git
   cd EmployeesTableReader
   ```
2. Open the solution file in Visual Studio.

3. Ensure you have the `EPPlus` library installed. You can install it via NuGet Package Manager:
   ```sh
   Install-Package EPPlus
   ```

4. Ensure your Excel file is located at `employees.xlsx` in the root of your project directory. The file should have the following structure:

   | Id | FirstName | LastName | Age | Profession | Salary | Diseases | IsOfficiallyEmployed |
   |----|-----------|----------|-----|------------|--------|----------|---------------------|
   | 1  | John      | Doe      | 30  | Engineer   | 5000   | Flu      | Yes                 |
   | 2  | Jane      | Smith    | 25  | Clerk      | 3000   | None     | No                  |
   | ...| ...       | ...      | ... | ...        | ...    | ...      | ...                 |

## Usage

1. Build and run the project in Visual Studio.

2. Use the `Program.cs` file to interact with the `EmployeesTable` class and test its functionalities.

Example usage in `Program.cs`:

```csharp
using System;
using System.Collections.Generic;

namespace EmployeesTableReader
{
    class Program
    {
        static void Main(string[] args)
        {
            EmployeesTable employeesTable = new EmployeesTable();

            // Fetch all employees
            List<Employee> allEmployees = employeesTable.GetAllEmployees();
            Console.WriteLine("All Employees:");
            foreach (var employee in allEmployees)
            {
                Console.WriteLine($"{employee.FirstName} {employee.LastName}, Profession: {employee.Profession}, Salary: {employee.Salary}");
            }

            // Fetch employees by profession
            Profession profession = Profession.Engineer;
            List<Employee> engineers = employeesTable.GetEmployeesByProfession(profession);
            Console.WriteLine($"\nEmployees with profession {profession}:");
            foreach (var engineer in engineers)
            {
                Console.WriteLine($"{engineer.FirstName} {engineer.LastName}, Salary: {engineer.Salary}");
            }

            // Calculate average salary
            double averageSalary = employeesTable.GetEmployeesAverageSalary();
            Console.WriteLine($"\nAverage Salary: {averageSalary}");

            // Fetch the highest paid employee
            Employee highestSalaryEmployee = employeesTable.GetEmployeeWithHighestSalary();
            Console.WriteLine($"\nEmployee with highest salary: {highestSalaryEmployee.FirstName} {highestSalaryEmployee.LastName}, Salary: {highestSalaryEmployee.Salary}");

            // Fetch unofficially employed employees
            List<Employee> unofficiallyEmployed = employeesTable.GetUnofficiallyEmployed();
            Console.WriteLine("\nUnofficially Employed Employees:");
            foreach (var employee in unofficiallyEmployed)
            {
                Console.WriteLine($"{employee.FirstName} {employee.LastName}");
            }

            // Fetch employees with salary above a threshold
            int salaryThreshold = 5000;
            List<Employee> highSalaryEmployees = employeesTable.GetEmployeesWithSalaryAbove(salaryThreshold);
            Console.WriteLine($"\nEmployees with salary above {salaryThreshold}:");
            foreach (var employee in highSalaryEmployees)
            {
                Console.WriteLine($"{employee.FirstName} {employee.LastName}, Salary: {employee.Salary}");
            }

            // Fetch employees younger than a specified age
            int ageThreshold = 30;
            List<Employee> youngEmployees = employeesTable.GetEmployeesByAge(ageThreshold, false);
            Console.WriteLine($"\nEmployees younger than {ageThreshold}:");
            foreach (var employee in youngEmployees)
            {
                Console.WriteLine($"{employee.FirstName} {employee.LastName}, Age: {employee.Age}");
            }

            // Fetch employees with a specific disease
            string disease = "Flu";
            List<Employee> employeesWithDisease = employeesTable.GetEmployeesByDisease(disease);
            Console.WriteLine($"\nEmployees with disease {disease}:");
            foreach (var employee in employeesWithDisease)
            {
                Console.WriteLine($"{employee.FirstName} {employee.LastName}, Diseases: {string.Join(", ", employee.Diseases)}");
            }

            // Add a new employee
            Employee newEmployee = new Employee
            {
                Id = 100,
                FirstName = "John",
                LastName = "Doe",
                Age = 25,
                Profession = Profession.Clerk,
                Salary = 3000,
                Diseases = new List<string> { "None" },
                IsOfficiallyEmployed = true
            };
            employeesTable.AddEmployee(newEmployee);
            Console.WriteLine($"\nAdded new employee: {newEmployee.FirstName} {newEmployee.LastName}");

            // Update an existing employee
            newEmployee.Salary = 3500;
            employeesTable.UpdateEmployee(newEmployee);
            Console.WriteLine($"\nUpdated employee salary: {newEmployee.FirstName} {newEmployee.LastName}, New Salary: {newEmployee.Salary}");

            // Delete an employee
            employeesTable.DeleteEmployee(newEmployee.Id);
            Console.WriteLine($"\nDeleted employee: {newEmployee.FirstName} {newEmployee.LastName}");
        }
    }
}
```

## Contributing

Contributions are welcome! If you have any suggestions or improvements, feel free to submit a pull request or open an issue.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

## Acknowledgments

- [EPPlus](https://github.com/EPPlusSoftware/EPPlus) library for Excel file operations.
- All contributors and open-source projects that helped in developing this application.


