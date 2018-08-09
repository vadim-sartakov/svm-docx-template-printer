package svm.msoffice.docx.printer.impl;

import java.io.FileInputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import static org.junit.Assert.*;
import svm.msoffice.docx.printer.utils.ExpectedValuesFactory;
import svm.msoffice.docx.printer.utils.Item;

public class TableParserTest {
    
    private DataHolder dataHolder;
    
    @Test
    public void testParse() throws Exception {
        
        XWPFDocument document = new XWPFDocument(
                new FileInputStream("src/test/resources/table/template.docx")
        );
                
        Item item = ExpectedValuesFactory.getItem();
        dataHolder = new DataHolder(item);
        
        assertEquals(ExpectedValuesFactory.getCorrectTableWithHeader(),
                new TableParser(dataHolder, document.getTableArray(0)).parse());
        assertEquals(ExpectedValuesFactory.getCorrectTableWithoutHeader(),
                new TableParser(dataHolder, document.getTableArray(2)).parse());
        item.getHistory().clear();
        assertEquals(ExpectedValuesFactory.getEmptyTable(),
                new TableParser(dataHolder, document.getTableArray(4)).parse());
                
    }
        
}
