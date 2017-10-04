/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.Map;
import svm.msoffice.docx.printer.utils.Item;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.Printer;
import svm.msoffice.docx.printer.utils.PrinterFactory;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;

/**
 *
 * @author sartakov
 */
public class ParserTest {
    
    private final Printer<Item> printer = PrinterFactory
            .getInstance("src/test/resources/parser/template.docx");
    
    @Test
    public void testParse() throws Exception {
               
        Map<Integer, Template> correctTemplates = ExpectedValuesFactory.getCorrectTemplates();
        
        XWPFParagraph paragraph = printer.getTemplateFile().getParagraphs().get(0);
        Map<Integer, Template> parsedTemplates = new Parser(printer, paragraph).parse();
            
        assertEquals(correctTemplates.size(), parsedTemplates.size());
        
        for (Map.Entry<Integer, Template> correctTemplate : correctTemplates.entrySet()) {
                        
            Template parsedTemplate = parsedTemplates.get(correctTemplate.getKey());
            assertNotNull(parsedTemplate);
            assertEquals(correctTemplate.getValue(), parsedTemplate);
            
        }
        
    }

}