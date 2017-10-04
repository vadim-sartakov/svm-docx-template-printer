/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.Map;
import static org.junit.Assert.*;
import org.junit.Test;
import svm.msoffice.docx.printer.utils.ExpectedFactory;

/**
 *
 * @author sartakov
 */
public class TemplateTest {

    private final Map<Integer, String> correctResults = ExpectedFactory.getRenderResults();
        
    @Test
    public void testRender() {
        
        Map<Integer, Template> correctTemplates = ExpectedFactory.getCorrectTemplates();
        correctTemplates.entrySet().forEach(entry -> {
            
            String actual = entry.getValue().render();
            String expected = correctResults.get(entry.getKey());
            
            assertEquals(expected, actual);
            
        });
        
    }
    
}
