package svm.msoffice.docx.printer.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class PrinterFactory {
    
    public static <T> Printer<T> getInstance(String input) {
        
        Item item = new Item();
        item.setName("Screwdriver");
        item.setDescription("Handy screwdriver");
        item.setPrice(new BigDecimal("12.6586"));
        item.setManufacturer("Some factory lmtd");
        item.setSerialNumber("15358-548");
        item.setReleaseDate(LocalDate.of(2012, 3, 25));
        item.setWeight("0.5 kg");
        item.setHeight("20 mm");
        item.setWidth("500 mm");
                
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
