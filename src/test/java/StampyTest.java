import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import stampy.Stampy;

import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalTime;
import java.util.Map;

public class StampyTest {
    private String input = "template.xlsx";
    private String output = "stampy_output.xlsx";
    private Stampy stampy;
    @Before
    public void setUp() throws Exception {
        stampy = Stampy.openTemplate(input);
    }

    @After
    public void tearDown() throws Exception {
    }

    @Test
    public void testMustache() {
        LocalTime start = LocalTime.now();
        stampy.executeTemplateMustaches(
                output,
                Map.of("rate", 10.4, "employees", 100_000, "profit", 245_001_000,
                        "words", new String[][]{
                                {"this", "is", "a", "sentence"},
                                {"so", "is", "this?"}
                        })
        );
        System.out.println(Duration.between(start, LocalTime.now()));
    }
}
