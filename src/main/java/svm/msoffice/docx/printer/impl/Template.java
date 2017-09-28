/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.text.DecimalFormat;
import java.time.format.DateTimeFormatter;
import java.time.temporal.Temporal;
import java.util.AbstractMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class Template {
    
    private final XWPFRun run;
    private String templateString;
    private String renderResult;
    public AbstractMap.SimpleEntry<String, String> format;
    public Integer width;
    public Map<String, Parameter> parameters;
    public Template enclosingTemplate;

    public Template(XWPFRun run, String templateString) {
        this.run = run;
        this.templateString = templateString;
        this.renderResult = templateString;
    }

    /**
     * Renders template into the specified run.
     */
    public void render() {

        if (enclosingTemplate != null) {

            String enclosingResult = templateString.replace(
                    enclosingTemplate.templateString,
                    enclosingTemplate.getRendered()
            );

            templateString = enclosingResult;
            renderResult = enclosingResult;

        }

        renderResult = getRendered();
        run.setText(renderResult, 0);

    }

    private String getRendered() {

        parameters.entrySet().forEach(parameterEntry ->
                renderParameter(parameterEntry));

        replaceFormatWithValues();
        renderResult = renderResult.replace("'", "");
        applyWidthIfPresent();

        return templateString.replace(templateString, renderResult);

    }

    private void renderParameter(Map.Entry<String, Parameter> parameterEntry) {
        String formattedValue = applyFormat(parameterEntry.getValue());
        renderResult = renderResult.replaceAll(Pattern.quote(parameterEntry.getKey()), formattedValue);
    }

    private String applyFormat(Parameter parameter) {

        // Leaving this apostrophes for decoupling formats and
        // values in one string if value is empty.
        String result = "''";
        Object value = parameter.getValue();

        if (value == null)
            return result;

        result = value.toString();

        if (format == null)
            return result;

        String formatValue = format.getValue();
        switch (format.getKey()) {

            case "date":
                Temporal convertedValue;
                if (value instanceof Temporal) {
                    convertedValue = (Temporal) value;
                    result = DateTimeFormatter.ofPattern(formatValue).format(convertedValue);
                }
                break;

            case "number":
                DecimalFormat decimalFormat = new DecimalFormat(formatValue);
                result = decimalFormat.format(value);
                break;

            default:

        }

        return result;

    }

    private void replaceFormatWithValues() {
        Matcher scopeMatcher = Printer.FORMAT_SCOPE_PATTERN.matcher(renderResult);
        if (scopeMatcher.find())
            renderResult = scopeMatcher.group(2);
    }

    private void applyWidthIfPresent() { 
        if (width != null)
            renderResult = StringUtils.center(renderResult, width);
    }
    
}
