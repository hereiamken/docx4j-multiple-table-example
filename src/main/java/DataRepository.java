import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

public class DataRepository {
    private final List<Employee> employees;

    DataRepository() {
        employees = new ArrayList<>();
        for(int i = 0;i< 50;i++){
            Employee em = new Employee();
            em.setNo(i+1);
            em.setName(getRandString());
            em.setAge(i+1);
            em.setDepartment("FPT_"+ getRandString());
            em.setPosition("DEV_"+getRandString());
            employees.add(em);
        }

    }

    private String getRandString() {
        List<String> strings = Arrays.asList("vimbo", "semimbee", "aginder", "neonu", "dynatude", "skamba", "avando", "premore", "conible", "polyil", "multixo", "parambo", "vicecy", "pixolium", "garent", "animoid", "conoodle", "dulia", "telender", "monic", "abafix", "amise", "leezzy", "albize", "postist", "amphinix");
        Random rand = new Random();
        String randomElement = strings.get(rand.nextInt(strings.size()));
        return randomElement;
    }

    List<Employee> getEmployees(){
        return employees;
    }
}
