package svm.msoffice.docx.printer.impl;

import java.io.FileInputStream;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;

public class TemplateParserTest {
    
    @Test
    public void testParse() throws Exception {
               
        Map<Integer, Template> correctTemplates = ExpectedValuesFactory.getCorrectTemplates();
        
        DataHolder dataHolder = new DataHolder(ExpectedValuesFactory.getItem());
        XWPFDocument document = new XWPFDocument(
                new FileInputStream("src/test/resources/parser/template.docx")
        );
        XWPFParagraph paragraph = document.getParagraphs().get(0);
        Map<Integer, Template> parsedTemplates = new TemplateParser(dataHolder, paragraph).parse();
            
        assertEquals(correctTemplates.size(), parsedTemplates.size());
        
        for (Map.Entry<Integer, Template> correctTemplate : correctTemplates.entrySet()) {
                        
            Template parsedTemplate = parsedTemplates.get(correctTemplate.getKey());
            assertNotNull(parsedTemplate);
            assertEquals(correctTemplate.getValue(), parsedTemplate);
            
        }
        
    }

}