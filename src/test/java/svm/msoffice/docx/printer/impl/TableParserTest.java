package svm.msoffice.docx.printer.impl;

import java.io.FileInputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;
import svm.msoffice.docx.printer.utils.Item;

/**
 *
 * @author sartakov
 */
public class TableParserTest {
    
    @Test
    public void testParse() throws Exception {
        
        XWPFDocument document = new XWPFDocument(
                new FileInputStream("src/test/resources/table/template.docx")
        );
        
        XWPFTable xwpfTable = document.getTableArray(0);
        
        Item item = ExpectedValuesFactory.getItem();
        DataHolder dataHolder = new DataHolder(item);
        TableParser tableParser = new TableParser(dataHolder, xwpfTable);
        Table actualTable = tableParser.parse();
        Table expectedTable = ExpectedValuesFactory.getCorrectTable();
        
        assertEquals(expectedTable, actualTable);
        
    }
    
}
