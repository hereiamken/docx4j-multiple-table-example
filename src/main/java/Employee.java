
public class Employee {
    private Integer no;
    private String name;
    private Integer age;
    private String position;
    private String department;

    public Employee() {
    }

    public Employee(Integer no, String name, Integer age, String position, String department) {
        this.no = no;
        this.name = name;
        this.age = age;
        this.position = position;
        this.department = department;
    }

    public Integer getNo() {
        return no;
    }

    public void setNo(Integer no) {
        this.no = no;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getPosition() {
        return position;
    }

    public void setPosition(String position) {
        this.position = position;
    }

    public String getDepartment() {
        return department;
    }

    public void setDepartment(String department) {
        this.department = department;
    }
}
