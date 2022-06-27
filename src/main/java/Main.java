import java.io.IOException;

public class Main {
    public static void main(String[] args) throws IOException {
        Docx2Adoc converter = new Docx2Adoc("C:\\Users\\Genius\\IdeaProjects\\ASCIIDoc\\Backend.docx",
                "C:\\Users\\Genius\\IdeaProjects\\ASCIIDoc\\result\\result.adoc");
        converter.convert();
    }

}
