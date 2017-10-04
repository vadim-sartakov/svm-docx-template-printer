package svm.msoffice.docx.printer;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class TemplateHolder implements AutoCloseable {
        
    private final XWPFDocument document;
    private final BufferedInputStream input;
        
    public TemplateHolder(String inputFilePath) {
    
        try {
            
            input = new BufferedInputStream(
                    new FileInputStream(
                            new File(inputFilePath)
                    )
            );
            document = new XWPFDocument(input);

        } catch(IOException e) {
            throw new RuntimeException(e);
        }
        
    }

    public void save() throws IOException {
        
        try (OutputStream output = new BufferedOutputStream(new FileOutputStream(
                new File("target/output.docx")))) {
            document.write(output);
        }
        
    }
    
    @Override
    public void close() throws Exception {
        input.close();
    }
    
    public XWPFDocument getDocument() {
        return document;
    }
    
}
