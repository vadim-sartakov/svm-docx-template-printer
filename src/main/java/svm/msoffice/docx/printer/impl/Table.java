package svm.msoffice.docx.printer.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author sartakov
 */
public class Table {
    
    private final XWPFTable xwpfTable;
    private final String name;
    private final DataHolder dataHolder;
    private final int templateRowIndex;
    private final List<Row> rows = new ArrayList<>();   
    
    private int tableRowIndex;
    private Row row;
    private XWPFTableRow xwpfRow;

    public Table(XWPFTable xwpfTable, String name, DataHolder dataHolder, int templateRowIndex) {
        this.xwpfTable = xwpfTable;
        this.name = name;
        this.dataHolder = dataHolder;
        this.templateRowIndex = templateRowIndex;
    }
    
    public void render() {
        populateTemplateRows();
        renderRows();        
    }
    
    private void populateTemplateRows() {
                        
        for (int index = templateRowIndex + 1; index <= rows.size(); index++) {
            
            XWPFTableRow newXwpfRow = index > xwpfTable.getRows().size() - 1 ?
                    xwpfTable.createRow() :
                    xwpfTable.insertNewTableRow(index);
            Utils.copyRow(xwpfTable.getRow(templateRowIndex), newXwpfRow);
                        
        }
            
    }
                    
    // TODO: move to parser
    private void renderRows() {
                     
        tableRowIndex = 0;
        for (Row currentRow : rows) {
                        
            dataHolder.putVariable("rowNumber", tableRowIndex + 1);
            row = currentRow;
            xwpfRow = xwpfTable.getRow(row.index);
            renderCells();
            
            tableRowIndex++;
            
        }
        
    }
    
    private void renderCells() {      
        for (Cell cell : row.cells)
            cell.render(xwpfRow.getCell(cell.index));        
    }
    
    
    
    public Row addRow(int index) {
        Row newRow = new Row(index);
        rows.add(newRow);
        return newRow;
    }
    
    public class Row {
        
        final int index;
        final List<Cell> cells = new ArrayList<>();
        
        public Row(int index) {
            this.index = index;
        }
        
        public Cell addCell(int index, Template template) {
            Cell cell = new Cell(index, template);
            cells.add(cell);
            return cell;
        }
                
    }
    
    public class Cell {
        
        final int index;
        final Map<Integer, Template> templates;
        
        public Cell(int index, Template template) {
            this.index = index;
            this.templates = new HashMap<>();
            this.templates.put(0, template);
        }
        
        public Cell(int index, Map<Integer, Template> templates) {
            this.index = index;
            this.templates = templates;
        }
        
        void render(XWPFTableCell xwpfTableCell) {
            xwpfTableCell.getParagraphs().forEach(paragraph -> {
                XWPFRunNormalizer.normalizeParameters(paragraph);
                insertIndexInParameters(paragraph);
                Template.renderTemplates(templates, paragraph);
            });
        }
        
        // TODO: move to parser
        private void insertIndexInParameters(XWPFParagraph paragraph) {
            for (XWPFRun currentRun : paragraph.getRuns()) {
                String runText = currentRun.getText(0);
                currentRun.setText(runText.replaceAll("\\{" + name,
                        "\\{" + name + "[" + tableRowIndex + "]"), 0);
            }
        }
        
    }
    
}
