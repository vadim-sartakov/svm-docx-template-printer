package svm.msoffice.docx.printer.impl;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author sartakov
 */
public class Utils {
    
    public static void copyRun(XWPFRun source, XWPFRun destination) {
        destination.getCTR().setRPr(source.getCTR().getRPr());
        destination.setText(source.getText(0), 0);
    }
    
    public static void copyParagraph(XWPFParagraph source, XWPFParagraph destination) {
        destination.getCTP().setPPr(source.getCTP().getPPr());
        source.getRuns().forEach(sourceRun -> {
            copyRun(sourceRun, destination.createRun());
        });
    }
    
}
