package svm.msoffice.docx.printer.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author sartakov
 */
public class Table {
    
    private final List<Row> rows = new ArrayList<>();   
    
    private XWPFTable xwpfTable;
    private int templateRowIndex;
    private Row row;
    private XWPFTableRow xwpfRow;
    
    public void render(XWPFTable xwpfTable) {
        
        this.xwpfTable = xwpfTable;
        
        if (rows.size() > 0) {
            this.templateRowIndex = rows.get(0).index;
        } else
            this.templateRowIndex = 0;
            
        populateTemplateRows();
        renderRows(); 
        
    }
    
    private void populateTemplateRows() {
                        
        for (int index = templateRowIndex + 1; index < rows.size() + templateRowIndex; index++) {
            
            XWPFTableRow newXwpfRow = index > xwpfTable.getRows().size() - 1 ?
                    xwpfTable.createRow() :
                    xwpfTable.insertNewTableRow(index);
            Utils.copyRow(xwpfTable.getRow(templateRowIndex), newXwpfRow);
                        
        }
            
    }
                    
    private void renderRows() {
                     
        for (Row currentRow : rows) {
                        
            row = currentRow;
            xwpfRow = xwpfTable.getRow(row.index);
            renderCells();

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
    
    public List<Row> getRows() {
        return rows;
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
        
        public List<Cell> getCells() {
            return cells;
        }
                
        public Cell addCell(int index, Map<Integer, Template> templates) {
            Cell cell = new Cell(index, templates);
            cells.add(cell);
            return cell;
        }

        @Override
        public int hashCode() {
            int hash = 7;
            hash = 17 * hash + this.index;
            hash = 17 * hash + Objects.hashCode(this.cells);
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
            final Row other = (Row) obj;
            if (this.index != other.index) {
                return false;
            }
            if (!Objects.equals(this.cells, other.cells)) {
                return false;
            }
            return true;
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
                Template.renderTemplates(templates, paragraph);
            });
        }

        public Map<Integer, Template> getTemplates() {
            return templates;
        }

        @Override
        public int hashCode() {
            int hash = 7;
            hash = 29 * hash + this.index;
            hash = 29 * hash + Objects.hashCode(this.templates);
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
            final Cell other = (Cell) obj;
            if (this.index != other.index) {
                return false;
            }
            if (!Objects.equals(this.templates, other.templates)) {
                return false;
            }
            return true;
        }
                
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 53 * hash + Objects.hashCode(this.rows);
        hash = 53 * hash + this.templateRowIndex;
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
        final Table other = (Table) obj;
        if (this.templateRowIndex != other.templateRowIndex) {
            return false;
        }
        if (!Objects.equals(this.rows, other.rows)) {
            return false;
        }
        return true;
    }
    
}
