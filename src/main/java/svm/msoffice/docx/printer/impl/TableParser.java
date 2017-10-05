package svm.msoffice.docx.printer.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author sartakov
 */
public class TableParser {
    
    private final static Pattern TABLE_CELL_PATTERN = Pattern.compile("\\{([^$\\.:\\}]+)\\.");
    
    private final List<Table> tables = new ArrayList<>();
    
    private int index;
    private List<Object> list;
    
    public TableParser(List<XWPFTable> tables) {
        
    }
    
    public List<Table> parse() {
        return tables;
    }
    
    /**
     * Any column in first section (before first dot)
     * should contain table name to iterate.
     */
    private void getNameAndValue() {
        
        index = 0;
        for (XWPFTableRow currentRow : table.getRows()) {
            
            row = currentRow;
            findTableName();
            if (!tableName.isEmpty()) {
                templateRowIndex = index;
                break;
            }
            
            index++;
            
        }
        
        if (tableName.isEmpty())
            return;
        
        try {
            list = (List) PropertyUtils.getProperty(object, tableName);
        } catch (Exception e) {
            LOGGER.warn("Failed to get list by name {}", tableName);
        }
        
        list = list == null ? new LinkedList<>() : list;
        
    }
    
    private void findTableName() {
                
        for (XWPFTableCell cell : row.getTableCells()) {
            
            String firstCellText = cell.getText();
            Matcher matcher = TABLE_CELL_PATTERN.matcher(firstCellText);
            if (matcher.find()) {
                tableName = matcher.group(1);
                break;
            } else
                tableName = "";
            
        }        
    
    }
     
}