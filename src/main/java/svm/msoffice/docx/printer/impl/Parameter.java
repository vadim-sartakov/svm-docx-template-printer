/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.Map;
import java.util.Objects;
import org.apache.commons.beanutils.PropertyUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class Parameter {
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    private final String property;
    private final Object value;

    public Parameter(Printer<?> printer, String property) {

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
            LOGGER.warn("Failed to get property {}", property);
        }
        
        Map<String, Converter> converters = printer.getConverters();
        if (converters != null) {
            Converter converter = printer.getConverters().get(property);
            retrievedValue = converter == null ? retrievedValue : converter.convert(retrievedValue);
        }
        
        value = retrievedValue;

    }

    public String getProperty() {
        return property;
    }
    
    public Object getValue() {
        return value;
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 53 * hash + Objects.hashCode(this.property);
        hash = 53 * hash + Objects.hashCode(this.value);
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
        final Parameter other = (Parameter) obj;
        if (!Objects.equals(this.property, other.property)) {
            return false;
        }
        if (!Objects.equals(this.value, other.value)) {
            return false;
        }
        return true;
    }
    
}
