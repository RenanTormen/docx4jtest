package dev.renantormen;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import javax.xml.bind.JAXBException;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 * Hello world!
 */
public final class App {
    private App() {
    }

    /**
     * Says hello to the world.
     * 
     * @param args The arguments of the program.
     */
    public static void main(String[] args) {
        try {
            WordprocessingMLPackage processor = WordprocessingMLPackage.load(new File("C:\\Projetos\\CN Soja.docx"));
            MainDocumentPart mainDocumentPart = processor.getMainDocumentPart();
            Map<String, String> varMap = new HashMap<>();
            varMap.put("REPLACE", "Esse valor foi substituido!");
            mainDocumentPart.variableReplace(varMap);
            SaveToZipFile saver = new SaveToZipFile(processor);
            saver.save("C:\\Projetos\\testeSalvo.docx");
            Docx4J.toPDF(processor, new FileOutputStream("C:\\Projetos\\teste.pdf"));
        } catch (JAXBException | Docx4JException | FileNotFoundException e) {
            e.printStackTrace();
        }
    }
}
