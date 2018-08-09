/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.junit.Test;
import static org.junit.Assert.*;

public class XWPFRunNormalizerTest {
                
    private String sourceParagraphText;
    XWPFParagraph paragraph;
    
    @Test
    public void testNormalizeRuns() throws Exception {
        normalizeAndSave();
        checkResults();        
    }
    
    private void normalizeAndSave() throws Exception {
        
        XWPFDocument document = new XWPFDocument(
                new FileInputStream("src/test/resources/normalizer/template.docx"));
            
        paragraph = document.getParagraphs().get(0);
        sourceParagraphText = paragraph.getText();
        XWPFRunNormalizer.normalizeParameters(paragraph);
        document.write(new FileOutputStream("target/output.docx"));
        
    }
    
    private void checkResults() throws Exception {
            
        XWPFDocument document = new XWPFDocument(
                new FileInputStream("target/output.docx"));
        
        paragraph = document.getParagraphs().get(0);
        assertEquals(sourceParagraphText, paragraph.getText());
        checkParameterRunConsistency();
        
    }
    
    private void checkParameterRunConsistency() {

        List<XWPFRun> runs = paragraph.getRuns();
        assertEquals("[{width: 20; number: \"0.00\"} ${price}]", runs.get(2).getText(0));
        assertEquals("${name}", runs.get(6).getText(0));
        assertEquals("${description}", runs.get(8).getText(0));
        assertEquals("[{width: 20} ${manufacturer} [{date: \"dd.MM.YYYY\"} ${releaseDate}]]", runs.get(12).getText(0));
        assertEquals("${weight}", runs.get(17).getText(0));
        assertEquals("${height}", runs.get(19).getText(0));
        assertEquals("${width}", runs.get(21).getText(0));
        
    }
    
}
