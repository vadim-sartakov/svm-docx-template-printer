/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.text.DecimalFormat;
import java.time.format.DateTimeFormatter;
import java.time.temporal.Temporal;
import java.util.AbstractMap.SimpleEntry;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.lang3.StringUtils;

/**
 *
 * @author sartakov
 */
public class Template {
    
    private String templateString;
    private String renderResult;
    public SimpleEntry<String, String> format;
    public Integer width;
    public Map<String, Object> parameterValues;
    public Template enclosingTemplate;

    public Template(String templateString) {
        this.templateString = templateString;
        this.renderResult = templateString;
    }

    /**
     * Renders template and returns render result.
     * @return rendered string.
     */
    public String render() {

        if (enclosingTemplate != null) {

            String enclosingResult = templateString.replace(
                    enclosingTemplate.templateString,
                    enclosingTemplate.getRendered()
            );

            templateString = enclosingResult;
            renderResult = enclosingResult;

        }

        renderResult = getRendered();
        
        return renderResult;

    }

    private String getRendered() {

        parameterValues.entrySet().forEach(parameterEntry ->
                renderParameter(parameterEntry));

        replaceFormatWithValues();
        renderResult = renderResult.replace("'", "");
        applyWidthIfPresent();

        return templateString.replace(templateString, renderResult);

    }

    private void renderParameter(Map.Entry<String, Object> parameterEntry) {
        String formattedValue = applyFormat(parameterEntry);
        renderResult = renderResult.replaceAll(Pattern.quote(parameterEntry.getKey()), formattedValue);
    }

    private String applyFormat(Map.Entry<String, Object> parameter) {

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
        Matcher scopeMatcher = TemplateParser.FORMAT_SCOPE_PATTERN.matcher(renderResult);
        if (scopeMatcher.find())
            renderResult = scopeMatcher.group(2);
    }

    private void applyWidthIfPresent() { 
        if (width != null)
            renderResult = StringUtils.center(renderResult, width);
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 13 * hash + Objects.hashCode(this.format);
        hash = 13 * hash + Objects.hashCode(this.width);
        hash = 13 * hash + Objects.hashCode(this.parameterValues);
        hash = 13 * hash + Objects.hashCode(this.enclosingTemplate);
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final Template other = (Template) obj;
        if (!Objects.equals(this.format, other.format)) {
            return false;
        }
        if (!Objects.equals(this.width, other.width)) {
            return false;
        }
        if (!Objects.equals(this.parameterValues, other.parameterValues)) {
            return false;
        }
        if (!Objects.equals(this.enclosingTemplate, other.enclosingTemplate)) {
            return false;
        }
        return true;
    }       
    
}
