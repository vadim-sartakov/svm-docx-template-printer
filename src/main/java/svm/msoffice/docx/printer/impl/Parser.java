package svm.msoffice.docx.printer.impl;

import java.util.AbstractMap.SimpleEntry;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import svm.msoffice.docx.printer.Printer;
import static svm.msoffice.docx.printer.Printer.FORMAT_PATTERN;
import static svm.msoffice.docx.printer.Printer.FORMAT_SCOPE_PATTERN;
import static svm.msoffice.docx.printer.Printer.PARAMETER_PATTERN;
import static svm.msoffice.docx.printer.Printer.PARAMETER_SCOPE_PATTERN;

/**
 *
 * @author sartakov
 */
public class Parser {
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    private final Printer<?> printer;
    private final XWPFParagraph paragraph;
    private final Map<Integer, Template> templates = new HashMap<>();
    private XWPFRun run;
    private int index;
    private Template template;
    
    public Parser(Printer<?> printer, XWPFParagraph paragraph) {
        this.printer = printer;
        this.paragraph = paragraph;
    }
    
    public Map<Integer, Template> parse() {
                
        index = 0;
        for (XWPFRun currentRun : paragraph.getRuns()) {
            
            run = currentRun;
            String runText = run.getText(0);
            
            template = parseFormattedParameters(runText);
            if (template != null)
                templates.put(index, template);
                        
            parseIndividualParameters(runText);
            
            index++;
            
        }      
        
        return templates;
        
    }
    
    private Template parseFormattedParameters(String templateString) {
                                         
        Matcher scopeMatcher = FORMAT_SCOPE_PATTERN.matcher(templateString);
        if (!scopeMatcher.find())
            return null;

        String formatString = scopeMatcher.group(1);
        String contentString = scopeMatcher.group(2);
        
        Template intermediateTemplate = new Template(run, scopeMatcher.group(0));
        parseFormats(intermediateTemplate, formatString);
        
        String contentWithParamsOnly = contentString.replaceAll(FORMAT_SCOPE_PATTERN.pattern(), "");      
        intermediateTemplate.parameterValues = parseParameters(contentWithParamsOnly);
        intermediateTemplate.enclosingTemplate = parseFormattedParameters(contentString);
        
        return intermediateTemplate;

    }
    
    private void parseFormats(Template intermediateTemplate, String formatString) {
        
        Matcher matcher = FORMAT_PATTERN.matcher(formatString);
        while (matcher.find()) {

            String name = matcher.group(1);
            String value = matcher.group(2);

            if (name.equals("width"))
                intermediateTemplate.width = Integer.parseInt(value);
            else
                intermediateTemplate.format = new SimpleEntry<>(name, value);

        }
        
    }
    
    private Map<String, Object> parseParameters(String templateString) {
        
        Map<String, Object> result = new HashMap<>();
        
        Matcher matcher = PARAMETER_PATTERN.matcher(templateString);
        while (matcher.find())
            result.put(matcher.group(0), getParameter(matcher.group(1)));
        
        return result;
        
    }
    
    private Object getParameter(String property) {
        
        Object value;
        Object variableValue = printer.getVariables().get(property);
        if (variableValue != null) {
            value = variableValue;
            return value;
        }

        Object retrievedValue = null;
        try {
            retrievedValue = PropertyUtils.getNestedProperty(printer.getObject(), property);            
        } catch (Exception e) {
            LOGGER.warn("Failed to get property {}", property);
        }
        
        Map<String, Converter> converters = printer.getConverters();
        if (converters != null) {
            Converter converter = printer.getConverters().get(property);
            retrievedValue = converter == null ? retrievedValue : converter.convert(retrievedValue);
        }
        
        value = retrievedValue;
        
        return value;
        
    }
    
    private void parseIndividualParameters(String templateString) {
                
        Matcher matcher = PARAMETER_SCOPE_PATTERN.matcher(templateString);
        if (!matcher.find())
            return;
        
        Map<String, Object> parameters = parseParameters(templateString);

        if (parameters.size() > 0) {
            template = new Template(run, matcher.group(0));
            template.parameterValues = parameters;
            templates.put(index, template);
        }
                
    }
    
}
