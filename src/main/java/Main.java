import com.sun.net.httpserver.Authenticator;

public class Main {
    public static void main(String[] args) throws Exception {

        String[][] a = {
                {"GROUP", "ИВБО-01-20"},
                {"FIO", "Белов Дмитрий Сергеевич"},
                {"PROGRAM", "Ежедневник"}
        };

        POI p = new POI();
        p.process(a,
                "P:\\files\\java\\replacerWord\\files\\test.docx",
                "P:\\files\\java\\replacerWord\\files\\output.docx");
    }
}