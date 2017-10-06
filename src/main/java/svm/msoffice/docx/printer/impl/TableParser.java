package svm.msoffice.docx.printer.impl;

import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import svm.msoffice.docx.printer.Printer;

/**
 *
 * @author sartakov
 */
public class TableParser {
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    private final static Pattern TABLE_CELL_PATTERN = Pattern.compile("\\{([^$\\.:\\}]+)\\.");
    
    private final DataHolder dataHolder;
    private final XWPFTable xwpfTable;
    
    private XWPFTableRow templateRow;
    private Table.Row row;
    private int templateRowIndex, rowIndex;
    private List<Object> list;
    private String tableName;
    private Table newTable;
    
    public TableParser(DataHolder dataHolder, XWPFTable xwpfTable) {
        this.dataHolder = dataHolder;
        this.xwpfTable = xwpfTable;
    }
    
    public Table parse() {
        getTableNameAndList();        
        return parseTable();
    }
    
    /**
     * Any column in first section (before first dot)
     * should contain table name to iterate.
     */
    private void getTableNameAndList() {
        
        int index = 0;
        for (XWPFTableRow currentRow : xwpfTable.getRows()) {
            
            templateRow = currentRow;
            findTableName();
            if (!tableName.isEmpty()) {
                templateRowIndex = index;
                templateRow = xwpfTable.getRow(index);
                break;
            }
            
            index++;
            
        }
        
        if (tableName.isEmpty())
            return;
        
        try {
            list = (List) PropertyUtils.getProperty(dataHolder.getObject(), tableName);
        } catch (Exception e) {
            LOGGER.warn("Failed to get list by name {}", tableName);
        }
        
        list = list == null ? new LinkedList<>() : list;
        
    }
    
    private void findTableName() {
                
        for (XWPFTableCell cell : templateRow.getTableCells()) {
            
            String firstCellText = cell.getText();
            Matcher matcher = TABLE_CELL_PATTERN.matcher(firstCellText);
            if (matcher.find()) {
                tableName = matcher.group(1);
                break;
            } else
                tableName = "";
            
        }        
    
    }
    
    private Table parseTable() {
        
        newTable = new Table();
        parseRows();
        
        return newTable;
                
    }
    
    private void parseRows() {
        
        if (list.isEmpty()) {
            
            row = newTable.addRow(templateRowIndex);
            dataHolder.putVariable("rowNumber", null);
            parseCells();
            
            return;
            
        }
        
        for (rowIndex = 0; rowIndex < list.size(); rowIndex++) {
            
            dataHolder.putVariable("rowNumber", rowIndex + 1);
            row = newTable.addRow(rowIndex + templateRowIndex);
                    
            parseCells();
            
        }
        
    }
        
    private void parseCells() {
        
        int cellIndex = 0;
        for (XWPFTableCell xwpfCell : templateRow.getTableCells()) {
            
            if (xwpfCell.getParagraphs().size() != 1)
                continue;
            
            XWPFParagraph paragraph = xwpfCell.getParagraphs().get(0);
            XWPFRunNormalizer.normalizeParameters(paragraph);
            insertIndexInParameters(paragraph);
            Map<Integer, Template> templates =
                    new TemplateParser(dataHolder, paragraph).parse();
            
            row.addCell(cellIndex, templates);
            removeIndexInParameters(paragraph);
            
            cellIndex++;
                    
        }
        
    }
    
    private void insertIndexInParameters(XWPFParagraph paragraph) {
        for (XWPFRun currentRun : paragraph.getRuns()) {
            String runText = currentRun.getText(0);
            currentRun.setText(runText.replaceAll("\\{" + tableName,
                    "\\{" + tableName + "[" + rowIndex + "]"), 0);
        }
    }
    
    private void removeIndexInParameters(XWPFParagraph paragraph) {
        for (XWPFRun currentRun : paragraph.getRuns()) {
            String runText = currentRun.getText(0);
            currentRun.setText(runText.replaceAll("\\[\\d\\]", ""), 0);
        }
    }
     
}