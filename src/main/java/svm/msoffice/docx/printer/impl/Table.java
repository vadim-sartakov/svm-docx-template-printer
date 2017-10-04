package svm.msoffice.docx.printer.impl;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import static svm.msoffice.docx.printer.Printer.TABLE_CELL_PATTERN;

/**
 *
 * @author sartakov
 */
public class Table {
    
    private final String name;
    private final List<Object> values;
    private final int startIndex;
    private final List<Template> rows;
    private final Map<String, Object> variables = new HashMap<>();
    
    private XWPFTable table;
    private int templateRowIndex, tableRowIndex;
    private XWPFTableRow row;
    private XWPFParagraph paragraph;

    public Table(String name, List<Object> values, int startIndex, List<Template> rows) {
        this.name = name;
        this.values = values;
        this.startIndex = startIndex;
        this.rows = rows;
    }
    
    public void render(XWPFTable table) {
        this.table = table;
        populateRows();
        renderRows();        
    }
    
    private void populateRows() {
                        
        for (int i = 0; i < values.size() - 1; i++) {
            XWPFTableRow newRow = table.createRow();
            copyRow(table.getRow(templateRowIndex), newRow);
        }
            
    }
    
    // Applying ctr copying leads to unexpected results. So, copying by hand.
    private void copyRow(XWPFTableRow source, XWPFTableRow destination) {
        
        int cellIndex = 0;
        for (XWPFTableCell sourceCell : source.getTableCells()) {
            
            XWPFTableCell destinationCell = destination.getCell(cellIndex);
            destinationCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            
            Utils.copyParagraph(sourceCell.getParagraphArray(0), destinationCell.getParagraphs().get(0));
            
            cellIndex++;
            
        }
        
    }
                
    private void renderRows() {
                     
        tableRowIndex = 0;
        for (XWPFTableRow currentRow : table.getRows()) {
                        
            variables.put("rowNumber", tableRowIndex + 1);
            
            // Bypassing headers
            if (!TABLE_CELL_PATTERN.matcher(currentRow.getCell(0).getText()).find())
                continue;
            
            row = currentRow;
            renderCells();
            
            tableRowIndex++;
            
        }
        
    }
    
    private void renderCells() {
                
        for (XWPFTableCell cell : row.getTableCells()) {
            
            for (XWPFParagraph currentParagraph : cell.getParagraphs()) {
                
                paragraph = currentParagraph;
                XWPFRunNormalizer.normalizeParameters(paragraph);
                
                insertIndexInParameters();
                renderTemplatesOfParagraph();
                
            }
        }
        
    }
    
    private void insertIndexInParameters() {
                
        for (XWPFRun currentRun : paragraph.getRuns()) {
            
            String runText = currentRun.getText(0);
            Matcher matcher = TABLE_CELL_PATTERN.matcher(runText);
            if (!matcher.find())
                continue;

            currentRun.setText(runText.replaceAll("\\{" + tableName,
                    "\\{" + tableName + "[" + tableRowIndex + "]"), 0);
            
        }
        
    }
    
}
