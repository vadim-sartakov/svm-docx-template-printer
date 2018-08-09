package svm.msoffice.docx.printer;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;
import svm.msoffice.docx.printer.utils.Item;

public class PrinterTest {
    
    private XWPFDocument expectedDocument, actualDocument;
    private XWPFTable expectedTable, actualTable;
    private XWPFTableRow expectedRow, actualRow;
    
    @Test
    public void testSomeMethod() throws Exception {
        
        Item item = ExpectedValuesFactory.getItem();
        Printer<Item> printer = new Printer<>(item,
                new FileInputStream("src/test/resources/printer/template.docx")
        );
        
        printer.print(new FileOutputStream("target/output.docx"));
        expectedDocument = new XWPFDocument(new FileInputStream("src/test/resources/printer/expected.docx"));
        actualDocument = new XWPFDocument(new FileInputStream("target/output.docx"));
        
        assertEquals(expectedDocument.getParagraphs().size(), expectedDocument.getParagraphs().size());
        assertEquals(expectedDocument.getTables().size(), expectedDocument.getTables().size());
        
        checkParagraphs();
        checkTables();
                
    }
    
    private void checkParagraphs() {
            
        int index = 0;
        for (XWPFParagraph expectedParagraph : expectedDocument.getParagraphs()) {
            assertEquals(expectedParagraph.getText(),
                    actualDocument.getParagraphs().get(index).getText());
            index++;
        }
            
    }
    
    private void checkTables() {
        
        expectedTable = expectedDocument.getTableArray(0);
        actualTable = actualDocument.getTableArray(0);
                
        checkRows();

    }
    
    private void checkRows() {
            
        assertEquals(expectedTable.getRows().size(), actualTable.getRows().size());
        for (int index = 0; index < expectedTable.getRows().size(); index++) {
            expectedRow = expectedTable.getRow(index);
            actualRow = actualTable.getRow(index);
            checkCells();
        }
    
    }
    
    private void checkCells() {
        
        for (int index = 0; index < expectedRow.getTableCells().size(); index++) {
         
            XWPFTableCell expectedCell = expectedRow.getCell(index);
            XWPFTableCell actualCell = actualRow.getCell(index);
            
            assertEquals(expectedCell.getParagraphs().size(), actualCell.getParagraphs().size());
            
            XWPFParagraph expectedParagraph = expectedCell.getParagraphs().get(0);
            XWPFParagraph actualParagraph = actualCell.getParagraphs().get(0);
            
            assertEquals(expectedParagraph.getText(), actualParagraph.getText());
            
        }
        
    }
    
}
