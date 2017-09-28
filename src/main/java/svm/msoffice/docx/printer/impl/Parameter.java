/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.Map;
import org.apache.commons.beanutils.PropertyUtils;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class Parameter {
    
    private final Printer printer;
    private final String property;
    private final Object value;

    public Parameter(Printer<?> printer, String property) {

        this.printer = printer;
        this.property = property;

        Object variableValue = printer.getVariables().get(property);
        if (variableValue != null) {
            value = variableValue;
            return;
        }

        Object retrievedValue = null;
        try {
            retrievedValue = PropertyUtils.getNestedProperty(printer.getObject(), property);            
        } catch (Exception e) {
            Printer.LOGGER.warn("Failed to get property {}", property);
        }
        
        Map<String, Converter> converters = printer.getConverters();
        if (converters != null) {
            Converter converter = printer.getConverters().get(property);
            retrievedValue = converter == null ? retrievedValue : converter.convert(retrievedValue);
        }
        
        value = retrievedValue;

    }

    public Printer getPrinter() {
        return printer;
    }

    public String getProperty() {
        return property;
    }
    
    public Object getValue() {
        return value;
    }
    
}
