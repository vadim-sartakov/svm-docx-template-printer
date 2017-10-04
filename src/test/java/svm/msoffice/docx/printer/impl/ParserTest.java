/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.AbstractMap.SimpleEntry;
import java.util.HashMap;
import java.util.Map;
import svm.msoffice.docx.printer.Item;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.Printer;
import svm.msoffice.docx.printer.PrinterFactory;

/**
 *
 * @author sartakov
 */
public class ParserTest {
    
    private final Printer<Item> printer = PrinterFactory.getInstance("src/test/resources/parser/template.docx");
    private final Map<Integer, Template> correctTemplates = new HashMap<>();
    
    @Test
    public void testParse() throws Exception {
               
        fillCorrectTemplates();
        
        XWPFParagraph paragraph = printer.getTemplateFile().getParagraphs().get(0);
        Map<Integer, Template> parsedTemplates = new Parser(printer, paragraph).parse();
            
        assertEquals(correctTemplates.size(), parsedTemplates.size());
        
        for (Map.Entry<Integer, Template> correctTemplate : correctTemplates.entrySet()) {
                        
            Template parsedTemplate = parsedTemplates.get(correctTemplate.getKey());
            assertNotNull(parsedTemplate);
            assertEquals(correctTemplate.getValue(), parsedTemplate);
            
        }
        
    }
    
    private void fillCorrectTemplates() {
        
        Template template, enclosingTemplate;
        
        template = new Template(null, "${name}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${name}", "Screwdriver");
        correctTemplates.put(2, template);
        
        template = new Template(null, "[{width: 20; number: \"0.00\"} ${price}]");
        template.format = new SimpleEntry<>("number", "0.00");
        template.width = 20;
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${price}", new BigDecimal("12.65"));
        correctTemplates.put(4, template);
                
        template = new Template(null, "${description}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${description}", "Handy screwdriver");
        correctTemplates.put(6, template);
        
        enclosingTemplate = new Template(null, "[{date: \"dd.MM.YYYY\"} ${releaseDate}]");
        enclosingTemplate.format = new SimpleEntry<>("date", "dd.MM.YYYY");
        enclosingTemplate.parameterValues = new HashMap<>();
        enclosingTemplate.parameterValues.put("${releaseDate}", LocalDate.of(2012, 3, 25));
        
        template = new Template(null, "[{width: 20} ${manufacturer} [{date: \"dd.MM.YYYY\"} ${releaseDate}]]");
        template.width = 20;
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${manufacturer}", "Some factory lmtd");
        template.enclosingTemplate = enclosingTemplate;
        correctTemplates.put(8, template);
        
        template = new Template(null, "${weight}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${weight}", "0.5 kg");
        correctTemplates.put(10, template);
        
        template = new Template(null, "${height}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${height}", "20 mm");
        correctTemplates.put(12, template);
        
        template = new Template(null, "${width}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${width}", "500 mm");
        correctTemplates.put(14, template);
                
    }
    
}
