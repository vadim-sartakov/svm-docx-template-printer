package svm.msoffice.docx.printer.impl;

import java.io.FileInputStream;
import java.util.AbstractMap.SimpleEntry;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import static org.junit.Assert.*;
import org.junit.Test;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;
import svm.msoffice.docx.printer.utils.Item;

/**
 *
 * @author sartakov
 */
public class TableTest {
    
    private Item item;
    private Table table;
    
    @Test
    public void testRender() throws Exception {
        
        XWPFDocument document = new XWPFDocument(
                new FileInputStream("src/test/resources/table/template.docx")
        );
        
        XWPFTable actualTable = document.getTableArray(0);
        XWPFTable expectedTable = document.getTableArray(1);
        
        item = ExpectedValuesFactory.getItem();
        DataHolder dataHolder = new DataHolder(item);
        
        table = new Table(actualTable, "history", dataHolder, 1);
        fillTableWithTemplates();
                
        table.render();
        
        assertEqualTables(expectedTable, actualTable);
        
    }
    
    private void fillTableWithTemplates() {
        
        int index = 1;
        for (Item.History historyItem : item.getHistory()) {
            
            Table.Row newRow = table.addRow(index);
            Template template;
            
            template = new Template("${rowNumber}", index);
            newRow.addCell(0, template);
            
            template = new Template("${history.date}");
            template.format = new SimpleEntry<>("date", "dd.MM.yyyy");
            template.parameterValues.put("${history.date}", historyItem.date);
            
            newRow.addCell(1, template);
            
            template = new Template("${history.price}");
            template.format = new SimpleEntry<>("number", "0.00");
            template.parameterValues.put("${history.price}", historyItem.price);
            
            newRow.addCell(2, template);
            newRow.addCell(3, new Template("${history.quantity}", historyItem.quantity));

            index++;
            
        }
        
    }
    
    private void assertEqualTables(XWPFTable expectedTable, XWPFTable actualTable) {
        
        assertEquals(expectedTable.getRows().size(), actualTable.getRows().size());
        
        for (int index = 0; index < expectedTable.getRows().size(); index++) {
            
            XWPFTableRow expectedRow = expectedTable.getRow(index);
            XWPFTableRow actualRow = actualTable.getRow(index);
                        
            assertEqualRows(expectedRow, actualRow);
                    
        }
        
    }
    
    private void assertEqualRows(XWPFTableRow expectedRow, XWPFTableRow actualRow) {
    
        assertEquals(expectedRow.getTableCells().size(), actualRow.getTableCells().size());
        
        for (int index = 0; index < expectedRow.getTableCells().size(); index++) {
            
            XWPFTableCell expectedCell = expectedRow.getTableCells().get(index);
            XWPFTableCell actualCell = actualRow.getTableCells().get(index);
                        
            assertEquals(expectedCell.getParagraphs().size(), actualCell.getParagraphs().size());
            
            XWPFParagraph expectedParagraph = expectedCell.getParagraphArray(0);
            XWPFParagraph actualParagraph = actualCell.getParagraphArray(0);
            
            assertEquals(expectedParagraph.getText(), actualParagraph.getText());
            
        }
    
    }
    
}
