package svm.msoffice.docx.printer.utils;

import java.io.FileInputStream;
import java.io.IOException;
import svm.msoffice.docx.printer.Printer;

public class PrinterFactory {
    
    public static Printer<Item> getInstance(String input) {
        
        Item item = ExpectedValuesFactory.getItem();
        Printer<Item> printer;
        
        try {
            printer = new Printer<>(
                    item,
                    new FileInputStream(input)
            );
        } catch(IOException e) {
            throw new RuntimeException(e);
        }

        return printer;
        
    }
        
    public static Printer<Item> getInstance() {
        return getInstance("src/test/resources/template.docx");
        
    }
    
}
