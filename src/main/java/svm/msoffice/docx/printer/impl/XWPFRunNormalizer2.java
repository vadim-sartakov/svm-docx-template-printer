package svm.msoffice.docx.printer.impl;

import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * MS Word arbitrary splits paragraph into separate runs.
 * As result, it makes impossible to parse parameters.
 * This class restores run consistency for further processing.
 * @author sartakov
 */
public class XWPFRunNormalizer2 {
    
    private final XWPFParagraph paragraph;
    private final List<XWPFRun> allRuns;
    private final Pattern pattern, startDistinguishPattern, endDistinguishPattern;   
    private final StringBuilder levelOneBuffer = new StringBuilder();
    
    private StringBuilder levelTwoBuffer;
    private XWPFRun firstRun, lastRun;
    private int firstIndex, lastIndex;
    
    public XWPFRunNormalizer2(XWPFParagraph paragraph, String regex) {
        this.paragraph = paragraph;
        this.allRuns = paragraph.getRuns();
        this.pattern = Pattern.compile(regex);
        this.startDistinguishPattern = Pattern.compile("(.+)(" + regex + ")");
        this.endDistinguishPattern = Pattern.compile("(" + regex + ")(.+)");
    }
    
    public void normalize() {
        if (!pattern.matcher(paragraph.getText()).find())
            return;
        runForwardLoop();
    }
    
    private void runForwardLoop() {
        
        for (lastIndex = 0; lastIndex < allRuns.size(); lastIndex++) {
                        
            lastRun = allRuns.get(lastIndex);
            levelOneBuffer.append(lastRun.getText(0));

            Matcher matcher = pattern.matcher(levelOneBuffer.toString());
            if (matcher.find()) {
                runBackwardLoop();
                levelOneBuffer.delete(matcher.start(), matcher.end());
            }
            
        }
        
    }
    
    private void runBackwardLoop() {
        
        levelTwoBuffer = new StringBuilder();
        for (firstIndex = lastIndex; firstIndex > 0; firstIndex--) {
            
            firstRun = allRuns.get(firstIndex);
            levelTwoBuffer.insert(0, firstRun.getText(0));
            
            if (pattern.matcher(levelTwoBuffer.toString()).find()) {
                collapse();
                distinguishStart();
                distinguishEnd();
                break;
            }
            
        }
        
    }
    
    private void collapse() {
        
        if (firstIndex == lastIndex)
            return;
        
        insertRun(firstRun, firstIndex, levelTwoBuffer.toString());
        for (int index = firstIndex + 1; index <= lastIndex + 1; index++)
            paragraph.removeRun(firstIndex + 1);
        
        int fragmentLength = lastIndex - firstIndex;
        lastIndex -= fragmentLength;
        
    }
    
    private void insertRun(XWPFRun styleSource, int index, String text) {
        XWPFRun newRun = paragraph.insertNewRun(index);
        Utils.copyRun(styleSource, newRun);
        newRun.setText(text, 0);
    }
            
    private void distinguishStart() {
       /* XWPFRun startRun = firstRun;
        Matcher matcher = startDistinguishPattern.matcher(startRun.getText(0));
        if (matcher.find()) {
            startRun.setText(matcher.group(1));
            addRun(lastIndex);
        }*/
    }
        
    private void distinguishEnd() {
       /* XWPFRun endRun = lastRun;
        Matcher matcher = endDistinguishPattern.matcher(endRun.getText(0));
        if (matcher.find()) {
            endRun.setText(matcher.group(2));
            addRun(firstIndex);
        }*/
    }
        
}
