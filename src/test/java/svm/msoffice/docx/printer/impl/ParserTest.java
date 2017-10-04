/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.AbstractMap;
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
        template.parameters = new HashMap<>();
        template.parameters.put("${name}", new Parameter(printer, "name"));
        correctTemplates.put(2, template);
        
        template = new Template(null, "[{width: 20; number: \"0.00\"} ${price}]");
        template.format = new AbstractMap.SimpleEntry<>("number", "0.00");
        template.width = 20;
        template.parameters = new HashMap<>();
        template.parameters.put("${price}", new Parameter(printer, "price"));
        correctTemplates.put(4, template);
                
        template = new Template(null, "${description}");
        template.parameters = new HashMap<>();
        template.parameters.put("${description}", new Parameter(printer, "description"));
        correctTemplates.put(6, template);
        
        enclosingTemplate = new Template(null, "[{date: \"dd.MM.YYYY\"} ${releaseDate}]");
        enclosingTemplate.format = new AbstractMap.SimpleEntry<>("date", "dd.MM.YYYY");
        enclosingTemplate.parameters = new HashMap<>();
        enclosingTemplate.parameters.put("${releaseDate}", new Parameter(printer, "releaseDate"));
        
        template = new Template(null, "[{width: 20} ${manufacturer} [{date: \"dd.MM.YYYY\"} ${releaseDate}]]");
        template.width = 20;
        template.parameters = new HashMap<>();
        template.parameters.put("${manufacturer}", new Parameter(printer, "manufacturer"));
        template.enclosingTemplate = enclosingTemplate;
        correctTemplates.put(8, template);
        
        template = new Template(null, "${weight}");
        template.parameters = new HashMap<>();
        template.parameters.put("${weight}", new Parameter(printer, "weight"));
        correctTemplates.put(10, template);
        
        template = new Template(null, "${height}");
        template.parameters = new HashMap<>();
        template.parameters.put("${height}", new Parameter(printer, "height"));
        correctTemplates.put(12, template);
        
        template = new Template(null, "${width}");
        template.parameters = new HashMap<>();
        template.parameters.put("${width}", new Parameter(printer, "width"));
        correctTemplates.put(14, template);
                
    }
    
}
