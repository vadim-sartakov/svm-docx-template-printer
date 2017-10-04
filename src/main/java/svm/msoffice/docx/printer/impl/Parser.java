package svm.msoffice.docx.printer.impl;

import java.util.AbstractMap;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
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
        intermediateTemplate.parameters = parseParameters(contentWithParamsOnly);
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
                intermediateTemplate.format = new AbstractMap.SimpleEntry<>(name, value);

        }
        
    }
    
    private Map<String, Parameter> parseParameters(String templateString) {
        
        Map<String, Parameter> result = new HashMap<>();
        
        Matcher matcher = PARAMETER_PATTERN.matcher(templateString);
        while (matcher.find())
            result.put(matcher.group(0), new Parameter(printer, matcher.group(1)));
        
        return result;
        
    }
    
    private void parseIndividualParameters(String templateString) {
                
        Matcher matcher = PARAMETER_SCOPE_PATTERN.matcher(templateString);
        if (!matcher.find())
            return;
        
        Map<String, Parameter> parameters = parseParameters(templateString);

        if (parameters.size() > 0) {
            template = new Template(run, matcher.group(0));
            template.parameters = parameters;
            templates.put(index, template);
        }
                
    }
    
}
