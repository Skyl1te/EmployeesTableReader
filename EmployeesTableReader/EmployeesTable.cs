using OfficeOpenXml;

namespace EmployeesTableReader;

public class EmployeesTable
{
    private const string _path = @"employees.xlsx";
    private ExcelWorksheet _worksheet;

    public EmployeesTable()
    {
        FileInfo file = new FileInfo(_path);
        ExcelPackage package = new ExcelPackage(file);

        _worksheet = package.Workbook.Worksheets[0];
    }

    private Employee GetEmployeeFromRow(int row)
    {
        Employee employee = new Employee
        {
            Id = _worksheet.Cells[row, 1].GetValue<int>(),
            FirstName = _worksheet.Cells[row, 2].GetValue<string>(),
            LastName = _worksheet.Cells[row, 3].GetValue<string>(),
            Age = _worksheet.Cells[row, 4].GetValue<int>(),
            Profession = Enum.Parse<Profession>(_worksheet.Cells[row, 5].GetValue<string>()),
            Salary = _worksheet.Cells[row, 6].GetValue<int>(),
            Diseases = _worksheet.Cells[row, 7].GetValue<string>().Split(";").ToList(),
            IsOfficiallyEmployed = _worksheet.Cells[row, 8].GetValue<string>() == "Yes"
        };

        return employee;
    }

    public List<Employee> GetAllEmployees()
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            Employee employee = GetEmployeeFromRow(row);

            employees.Add(employee);
        }

        return employees;
    }

    public Employee? GetEmployeeById(int id)
    {
        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            if (_worksheet.Cells[row, 1].GetValue<int>() == id)
            {
                return GetEmployeeFromRow(row);
            }
        }

        return null;
    }

    public List<Employee> GetUnofficiallyEmployed()
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            if (_worksheet.Cells[row, 8].GetValue<string>() == "No")
            {
                employees.Add(GetEmployeeFromRow(row));
            }
        }

        return employees;
    }

    public List<Employee> GetEmployeesByProfession(Profession profession)
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            if (Enum.Parse<Profession>(_worksheet.Cells[row, 5].GetValue<string>()) == profession)
            {
                employees.Add(GetEmployeeFromRow(row));
            }
        }

        return employees;
    }

    public double GetEmployeesAverageSalary()
    {
        int cnt = 0;
        int sum = 0;
        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            cnt++;
            sum += _worksheet.Cells[row, 6].GetValue<int>();
        }
                
        return (double)(sum / cnt);
    }

    public Employee GetEmployeeWithHighestSalary()
    {
        int tmpSalary = 0;
        int tmpSalaryRow = -1;

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            int salary = _worksheet.Cells[row, 6].GetValue<int>();
            if (salary > tmpSalary)
            {
                tmpSalary = salary;
                tmpSalaryRow = row;
            }
        }

        if (tmpSalaryRow == -1)
        {
            throw new InvalidOperationException("No employees found");
        }

        return GetEmployeeFromRow(tmpSalaryRow);
    }

    public List<Employee> GetEmployeesWithSalaryAbove(int threshold)
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            int salary = _worksheet.Cells[row, 6].GetValue<int>();
            if (salary > threshold)
            {
                employees.Add(GetEmployeeFromRow(row));
            }
        }

        return employees;
    }

    public List<Employee> GetEmployeesByAge(int age)
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            int employeeAge = _worksheet.Cells[row, 4].GetValue<int>();
            if (employeeAge == age)
            {
                employees.Add(GetEmployeeFromRow(row));
            }
        }

        return employees;
    }

    public List<Employee> GetEmployeesByDisease(string disease)
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            string diseases = _worksheet.Cells[row, 7].GetValue<string>();
            if (diseases.Split(";").Contains(disease))
            {
                employees.Add(GetEmployeeFromRow(row));
            }
        }

        return employees;
    }

    public List<Employee> GetOfficiallyEmployed()
    {
        List<Employee> employees = new List<Employee>();

        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            if (_worksheet.Cells[row, 8].GetValue<string>() == "Yes")
            {
                employees.Add(GetEmployeeFromRow(row));
            }
        }

        return employees;
    }

    public void AddEmployee(Employee employee)
    {
        int newRow = _worksheet.Dimension.Rows + 1;
        _worksheet.Cells[newRow, 1].Value = employee.Id;
        _worksheet.Cells[newRow, 2].Value = employee.FirstName;
        _worksheet.Cells[newRow, 3].Value = employee.LastName;
        _worksheet.Cells[newRow, 4].Value = employee.Age;
        _worksheet.Cells[newRow, 5].Value = employee.Profession.ToString();
        _worksheet.Cells[newRow, 6].Value = employee.Salary;
        _worksheet.Cells[newRow, 7].Value = string.Join(";", employee.Diseases);
        _worksheet.Cells[newRow, 8].Value = employee.IsOfficiallyEmployed ? "Yes" : "No";
    }

    public void UpdateEmployee(Employee updatedEmployee)
    {
        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            if (_worksheet.Cells[row, 1].GetValue<int>() == updatedEmployee.Id)
            {
                _worksheet.Cells[row, 2].Value = updatedEmployee.FirstName;
                _worksheet.Cells[row, 3].Value = updatedEmployee.LastName;
                _worksheet.Cells[row, 4].Value = updatedEmployee.Age;
                _worksheet.Cells[row, 5].Value = updatedEmployee.Profession.ToString();
                _worksheet.Cells[row, 6].Value = updatedEmployee.Salary;
                _worksheet.Cells[row, 7].Value = string.Join(";", updatedEmployee.Diseases);
                _worksheet.Cells[row, 8].Value = updatedEmployee.IsOfficiallyEmployed ? "Yes" : "No";
                break;
            }
        }
    }

    public void DeleteEmployee(int id)
    {
        for (int row = 2; row <= _worksheet.Dimension.Rows; row++)
        {
            if (_worksheet.Cells[row, 1].GetValue<int>() == id)
            {
                _worksheet.DeleteRow(row);
                break;
            }
        }
    }



}
