package svm.msoffice.docx.printer.utils;

import java.io.FileInputStream;
import java.io.IOException;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class PrinterFactory {
    
    public static <T> Printer<T> getInstance(String input) {
        
        Item item = ExpectedValuesFactory.getItem();
        Printer<T> printer;
        
        try {
            printer = new Printer(
                    item,
                    new FileInputStream(input)
            );
        } catch(IOException e) {
            throw new RuntimeException(e);
        }

        return printer;
        
    }
        
    public static <T> Printer<T> getInstance() {
        return getInstance("src/test/resources/template.docx");
        
    }
    
}
