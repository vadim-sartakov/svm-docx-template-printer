package svm.msoffice.docx.printer.impl;

import java.util.Map;
import static org.junit.Assert.*;
import org.junit.Test;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;

public class TemplateTest {

    private final Map<Integer, String> correctResults = ExpectedValuesFactory.getRenderResults();
        
    @Test
    public void testRender() {
        
        Map<Integer, Template> correctTemplates = ExpectedValuesFactory.getCorrectTemplates();
        correctTemplates.entrySet().forEach(entry -> {
            
            String actual = entry.getValue().render();
            String expected = correctResults.get(entry.getKey());
            
            assertEquals(expected, actual);
            
        });
        
    }
    
}
