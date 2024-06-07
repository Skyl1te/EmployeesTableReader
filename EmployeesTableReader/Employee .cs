namespace EmployeesTableReader;
public class Employee
{
    public int Id { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public int Age { get; set; }
    public int Salary { get; set;}
    public List<string> Diseases { get; set; }
    public bool IsOfficiallyEmployed { get; set; }
    public Profession Profession { get; set; }

}

public enum Profession
{
    BackendDeveloper,
    FrontendDeveloper,
    DevOpsEngineer,
    SystemAdministrator,
    QAEnginner,
    HR,
    Designer
}