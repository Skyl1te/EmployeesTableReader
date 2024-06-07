using EmployeesTableReader;
using OfficeOpenXml;
using System.ComponentModel.Design;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


EmployeesTable employeesTable = new EmployeesTable();

// Getting all employees
List<Employee> allEmployees = employeesTable.GetAllEmployees();
Console.WriteLine("All Employees:");
foreach (var employee in allEmployees)
{
    Console.WriteLine($"{employee.FirstName} {employee.LastName}, Profession: {employee.Profession}, Salary: {employee.Salary}");
}
Console.WriteLine("_______________________");

// Getting employees by profession
Profession profession = Profession.QAEnginner;
List<Employee> qaEngineers = employeesTable.GetEmployeesByProfession(profession);
Console.WriteLine($"\nEmployees with profession {profession}:");
foreach (var engineer in qaEngineers)
{
    Console.WriteLine($"{engineer.FirstName} {engineer.LastName}, Salary: {engineer.Salary}");
}
Console.WriteLine("_______________________");

// Getting the average salary of all employees
double averageSalary = employeesTable.GetEmployeesAverageSalary();
Console.WriteLine($"\nAverage Salary: {averageSalary}");
Console.WriteLine("_______________________");

// Getting the employee with the highest salary
Employee highestSalaryEmployee = employeesTable.GetEmployeeWithHighestSalary();
Console.WriteLine($"\nEmployee with highest salary: {highestSalaryEmployee.FirstName} {highestSalaryEmployee.LastName}, Salary: {highestSalaryEmployee.Salary}");
Console.WriteLine("_______________________");

// Getting unofficially employed employees
List<Employee> unofficiallyEmployed = employeesTable.GetUnofficiallyEmployed();
Console.WriteLine("\nUnofficially Employed Employees:");
foreach (var employee in unofficiallyEmployed)
{
    Console.WriteLine($"{employee.FirstName} {employee.LastName}");
}
Console.WriteLine("_______________________");

// Getting employees with salaries above a certain value
int salaryThreshold = 5000;
List<Employee> highSalaryEmployees = employeesTable.GetEmployeesWithSalaryAbove(salaryThreshold);
Console.WriteLine($"\nEmployees with salary above {salaryThreshold}:");
foreach (var employee in highSalaryEmployees)
{
    Console.WriteLine($"{employee.FirstName} {employee.LastName}, Salary: {employee.Salary}");
}

// Getting employees 30 years of age
int ageThreshold = 30;
List<Employee> youngEmployees = employeesTable.GetEmployeesByAge(ageThreshold);
Console.WriteLine($"\nEmployees younger than {ageThreshold}:");
foreach (var employee in youngEmployees)
{
    Console.WriteLine($"{employee.FirstName} {employee.LastName}, Age: {employee.Age}");
}

// Receiving employees with a certain disease
string disease = "Headache";
List<Employee> employeesWithDisease = employeesTable.GetEmployeesByDisease(disease);
Console.WriteLine($"\nEmployees with disease {disease}:");
foreach (var employee in employeesWithDisease)
{
    Console.WriteLine($"{employee.FirstName} {employee.LastName}, Diseases: {string.Join(", ", employee.Diseases)}");
}

// Adding a new employee
Employee newEmployee = new Employee
{
    Id = 100,
    FirstName = "John",
    LastName = "Doe",
    Age = 25,
    Profession = Profession.HR,
    Salary = 3000,
    Diseases = new List<string> { "None" },
    IsOfficiallyEmployed = true
};
employeesTable.AddEmployee(newEmployee);
Console.WriteLine($"\nAdded new employee: {newEmployee.FirstName} {newEmployee.LastName}");

// Employee update
newEmployee.Salary = 3500;
employeesTable.UpdateEmployee(newEmployee);
Console.WriteLine($"\nUpdated employee salary: {newEmployee.FirstName} {newEmployee.LastName}, New Salary: {newEmployee.Salary}");

// Removing an employee
employeesTable.DeleteEmployee(newEmployee.Id);
Console.WriteLine($"\nDeleted employee: {newEmployee.FirstName} {newEmployee.LastName}");
        