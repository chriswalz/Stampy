import stampy.Stampy;

import java.util.Map;

public class Main {
    // Java 9
    // Templates for Reports
    // POI Name Manager
    // Reflection
    public static void main(String[] args) {
        generateReport();
    }

    private static void generateReport() {
        System.out.println("--open--");
        Stampy stampy = Stampy.openTemplate("template.xlsx");

        System.out.println("--output report based on Map<String, Object>");
        stampy.executeTemplateMustaches(
                "stampy_output.xlsx",
                Map.of("rate", 10.4, "employees", 100_000, "profit", 245_001_000,
                        "words", new String[][]{
                                {"this", "is", "a", "sentence"},
                                {"so", "is", "this?"}
                        })
        );
    }
}
