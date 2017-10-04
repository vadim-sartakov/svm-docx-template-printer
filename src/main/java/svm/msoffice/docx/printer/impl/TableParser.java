package svm.msoffice.docx.printer.impl;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFTable;

/**
 *
 * @author sartakov
 */
public class TableParser {
    
    private final Map<Integer, Table> tables = new HashMap<>();
    
    public TableParser(List<XWPFTable> tables) {
        
    }
    
    public Map<Integer, Table> parse() {
        return tables;
    }
     
}