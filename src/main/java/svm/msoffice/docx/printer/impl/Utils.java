package svm.msoffice.docx.printer.impl;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class Utils {
    
    // Applying ctr copying leads to unexpected results. So, copying by hand.
    public static void copyRow(XWPFTableRow source, XWPFTableRow destination) {
        
        int cellIndex = 0;
        for (XWPFTableCell sourceCell : source.getTableCells()) {
            
            XWPFTableCell destinationCell = destination.getCell(cellIndex);
            destinationCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            
            Utils.copyParagraph(sourceCell.getParagraphArray(0), destinationCell.getParagraphs().get(0));
            
            cellIndex++;
            
        }
        
    }
    
    public static void copyParagraph(XWPFParagraph source, XWPFParagraph destination) {
        destination.getCTP().setPPr(source.getCTP().getPPr());
        source.getRuns().forEach(sourceRun -> {
            copyRun(sourceRun, destination.createRun());
        });
    }
    
    public static void copyRun(XWPFRun source, XWPFRun destination) {
        destination.getCTR().setRPr(source.getCTR().getRPr());
        destination.setText(source.getText(0), 0);
    }
        
}
