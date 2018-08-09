package svm.msoffice.docx.printer.impl;

import java.io.FileInputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import static org.junit.Assert.*;
import org.junit.Test;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;

public class TableTest {
    
    private XWPFDocument document;
    private Table table;
    
    @Test
    public void testRender() throws Exception {
        
        document = new XWPFDocument(
                new FileInputStream("src/test/resources/table/template.docx")
        );
        
        testTables(1, 1, 0, ExpectedValuesFactory.getCorrectTableWithHeader());
        testTables(0, 3, 2, ExpectedValuesFactory.getCorrectTableWithoutHeader());
        testTables(1, 5, 4, ExpectedValuesFactory.getEmptyTable());
        
    }
    
    private void testTables(int templateRowIndex, int expectedIndex, int actualIndex, Table actualTable) {
        
        XWPFTable actualXwpfTable = document.getTableArray(actualIndex);
        actualXwpfTable.getRow(templateRowIndex).getTableCells().forEach(cell ->
                XWPFRunNormalizer.normalizeParameters(cell.getParagraphs().get(0))
        );
        XWPFTable expectedXwpfTable = document.getTableArray(expectedIndex);
        actualTable.render(actualXwpfTable);
        assertEqualTables(expectedXwpfTable, actualXwpfTable);
        
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
