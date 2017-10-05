package svm.msoffice.docx.printer.impl;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * Contains shared data.
 * Holds object, converters and variables.
 * @author sartakov
 */
public class DataHolder {
        
    private final Object object;
    private final Map<String, Converter> converters;
    private final Map<String, Object> variables = new HashMap<>();

    public DataHolder(Object object) {
        this(object, null);
    }
    
    public DataHolder(Object object, Map<String, Converter> converters) {
        this.object = object;
        this.converters = converters == null ? new HashMap<>() : converters;
    }
        
    public synchronized Object getVariable(String key) {
        return variables.get(key);
    }
    
    public synchronized void putVariable(String key, Object value) {
        variables.put(key, value);
    }

    public Map<String, Converter> getConverters() {
        return Collections.unmodifiableMap(converters);
    }
    
    public Object getObject() {
        return object;
    }
            
}
