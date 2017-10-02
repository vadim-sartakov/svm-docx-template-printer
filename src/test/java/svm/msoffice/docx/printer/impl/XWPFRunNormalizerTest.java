/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.TemplateHolder;

/**
 *
 * @author sartakov
 */
public class XWPFRunNormalizerTest {
                
    private String sourceParagraphText;
    XWPFParagraph paragraph;
    
    @Test
    public void testNormalizeRuns() {
        normalizeAndSave();
        checkResults();        
    }
    
    private void normalizeAndSave() {
        
        try (TemplateHolder templateHolder = new TemplateHolder()) {
            
            paragraph = templateHolder.getDocument().getParagraphs().get(0);
            sourceParagraphText = paragraph.getText();
            
            new XWPFRunNormalizer(paragraph, "\\$\\{.+\\}").normalize();
            new XWPFRunNormalizer(
                    paragraph,
                    "\\[\\{.+\\]", 
                    textFragment -> 
                            StringUtils.countMatches(textFragment, "[") ==
                            StringUtils.countMatches(textFragment, "]"),
                    "[^\\[]+"
            ).normalize();

            templateHolder.save();
            
        } catch(Exception e) {
            fail(e.getMessage());
        }
        
    }
    
    private void checkResults() {
            
        try (TemplateHolder templateHolder = new TemplateHolder("target/output.docx")) {
            
            paragraph = templateHolder.getDocument().getParagraphs().get(0);
            assertEquals(sourceParagraphText, paragraph.getText());
            checkParameterRunConsistency();
            
        } catch(Exception e) {
            fail(e.getMessage());
        }
        
    }
    
    private void checkParameterRunConsistency() {

        List<XWPFRun> runs = paragraph.getRuns();
        assertEquals(runs.get(2).getText(0), "[{width: 20; number: \"0.00\"} ${price}]");
        assertEquals(runs.get(6).getText(0), "${name}");
        assertEquals(runs.get(8).getText(0), "${description}");
        assertEquals(runs.get(12).getText(0), "[{width: 20} ${manufacturer} [{date: \"dd.MM.YYYY\"} ${releaseDate}]]");
        
    }
    
}
