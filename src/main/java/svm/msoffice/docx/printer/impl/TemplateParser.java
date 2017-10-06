package svm.msoffice.docx.printer.impl;

import svm.msoffice.docx.printer.Converter;
import java.util.AbstractMap.SimpleEntry;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class TemplateParser {
    
    /**
     * Format and parameters
     */
    public final static Pattern FORMAT_SCOPE_PATTERN = Pattern.compile("\\[(\\{[^$]*\\}) +(.+)\\]");
    /**
     * Parameters only, without format scope
     */
    public final static Pattern PARAMETER_SCOPE_PATTERN = Pattern.compile("(\\$\\{([\\w.\\[\\]]+)\\})(?!.*\\])");
    /**
     * Format - value
     */
    public final static Pattern FORMAT_PATTERN = Pattern.compile("([\\w-]+): *\"?([\\wА-Яа-я .']+)\"?");
    /**
     * Parameter - content
     */
    public final static Pattern PARAMETER_PATTERN = Pattern.compile("\\$\\{([\\w.\\[\\]]+)\\}");
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    private final DataHolder dataHolder;
    private final XWPFParagraph paragraph;
    private final Map<Integer, Template> templates = new HashMap<>();
    private XWPFRun run;
    private int index;
    private Template template;
    
    public TemplateParser(DataHolder dataHolder, XWPFParagraph paragraph) {
        this.dataHolder = dataHolder;
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
        
        Template intermediateTemplate = new Template(scopeMatcher.group(0));
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
        Object variableValue = dataHolder.getVariable(property);
        if (variableValue != null) {
            value = variableValue;
            return value;
        }

        Object retrievedValue = null;
        try {
            retrievedValue = PropertyUtils.getNestedProperty(dataHolder.getObject(), property);            
        } catch (Exception e) {
            LOGGER.warn("Failed to get property {}", property);
        }
        
        Map<String, Converter> converters = dataHolder.getConverters();
        if (converters != null) {
            Converter converter = dataHolder.getConverters().get(property);
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
            template = new Template(matcher.group(0));
            template.parameterValues = parameters;
            templates.put(index, template);
        }
                
    }
    
}
