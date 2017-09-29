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
        
    private XWPFRun firstLevelRun, secondLevelRun;
    private int firstLevelIndex, secondLevelIndex;
    
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
        
        StringBuilder levelOneBuffer = new StringBuilder();
        for (firstLevelIndex = 0; firstLevelIndex < allRuns.size(); firstLevelIndex++) {
                        
            firstLevelRun = allRuns.get(firstLevelIndex);
            levelOneBuffer.append(firstLevelRun.getText(0));
            
            if (pattern.matcher(levelOneBuffer.toString()).find()) {
                runSecondLevelLoop();
                levelOneBuffer = new StringBuilder();
            }
            
        }
        
    }
    
    private void runSecondLevelLoop() {
        
        StringBuilder levelTwoBuffer = new StringBuilder();
        for (secondLevelIndex = firstLevelIndex; secondLevelIndex > 0; secondLevelIndex--) {
            
            secondLevelRun = allRuns.get(secondLevelIndex);
            levelTwoBuffer.insert(0, secondLevelRun.getText(0));
            
            if (pattern.matcher(levelTwoBuffer.toString()).find()) {
                distinguishPattern();
                break;
            }
            
        }
        
    }
    
    private void distinguishPattern() {
        distinguishStart();
        distinguishEnd();
        collapse();
    }
        
    private void distinguishStart() {
        XWPFRun startRun = secondLevelRun;
        Matcher matcher = startDistinguishPattern.matcher(startRun.getText(0));
        if (matcher.find()) {
            startRun.setText(matcher.group(1));
            addRun(firstLevelIndex);
        }
    }
    
    private void addRun(int index) {
        
    }
    
    private void distinguishEnd() {
        XWPFRun endRun = firstLevelRun;
        Matcher matcher = endDistinguishPattern.matcher(endRun.getText(0));
        if (matcher.find()) {
            endRun.setText(matcher.group(2));
            addRun(secondLevelIndex);
        }
    }
    
    private void collapse() {
        
    }
    
}
