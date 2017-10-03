package svm.msoffice.docx.printer.impl;

import java.util.List;
import java.util.function.Predicate;
import svm.misc.regex.Matcher;
import svm.misc.regex.Pattern;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * MS Word arbitrary splits paragraph into separate runs.
 * As result, it makes impossible to parse parameters.
 * This class restores run consistency for further processing.
 * @author sartakov
 */
public class XWPFRunNormalizer {
    
    private final XWPFParagraph paragraph;
    private final Predicate<String> condition;
    private final List<XWPFRun> allRuns;
    private final Pattern pattern, startSplitPattern, endSplitPattern;   

    private XWPFRun firstRun, lastRun;
    private int firstIndex, lastIndex;
    
    public XWPFRunNormalizer(XWPFParagraph paragraph, String regex) {
        this(paragraph, regex, null, ".+");
    }
    
    public XWPFRunNormalizer(XWPFParagraph paragraph, String regex, Predicate<String> condition) {
        this(paragraph, regex, condition, ".+");
    }
        
    /**
     * 
     * @param paragraph
     * @param regex
     * @param condition - some extra condition to test if it's not enough for regex.
     * @param boundSplitRestRegex - regex for capturing rest in the bound splitting procedure.
     */
    public XWPFRunNormalizer(XWPFParagraph paragraph,
            String regex,
            Predicate<String> condition,
            String boundSplitRestRegex) {
        this.paragraph = paragraph;
        this.condition = condition == null ? textFragment -> true : condition;
        this.allRuns = paragraph.getRuns();
        this.pattern = Pattern.compile(regex);
        this.startSplitPattern = Pattern.compile("(" + boundSplitRestRegex + ")((" + regex + ")(" + boundSplitRestRegex + ")?)");
        this.endSplitPattern = Pattern.compile("(" + regex + ")(" + boundSplitRestRegex + ")");
    }
    
    public void normalize() {
        if (!pattern.matcher(paragraph.getText()).find())
            return;
        runForwardLoop();
    }
    
    private void runForwardLoop() {
        
        StringBuilder levelOneBuffer = new StringBuilder();
        for (lastIndex = 0; lastIndex < allRuns.size(); lastIndex++) {
                        
            lastRun = allRuns.get(lastIndex);
            levelOneBuffer.append(lastRun.getText(0));
            
            String bufferOneResult = levelOneBuffer.toString();
            
            Matcher matcher = pattern.matcher(bufferOneResult);
            if (matcher.find() && condition.test(bufferOneResult)) {
                runBackwardLoop(levelOneBuffer.substring(matcher.start(), matcher.end()));
                levelOneBuffer = new StringBuilder();
            }
            
        }
        
    }
    
    private void runBackwardLoop(String matchedString) {
        
        StringBuilder levelTwoBuffer = new StringBuilder();
        for (firstIndex = lastIndex; firstIndex > 0; firstIndex--) {
            
            firstRun = allRuns.get(firstIndex);
            levelTwoBuffer.insert(0, firstRun.getText(0));
            
            String resultText = levelTwoBuffer.toString();
            
            if (resultText.contains(matchedString)) {
                
                collapse(resultText);
                splitRepeat(resultText);
                splitBound(lastIndex, startSplitPattern);
                splitBound(lastIndex, endSplitPattern);
                
                break;
                
            }
            
        }
        
    }
    
    private void collapse(String resultText) {
        
        if (firstIndex == lastIndex)
            return;
        
        insertRun(firstRun, firstIndex, resultText);
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
    
    /**
     * In case if run contains repeated pattern.
     */
    private void splitRepeat(String resultText) {
        
        StringBuilder stringBuilder = new StringBuilder(resultText);
        Matcher matcher = pattern.matcher(resultText);
        
        int totalMatchCount = getMatchCountAndReset(matcher);
        if (totalMatchCount == 1)
            return;
        
        int matchCount = 1;
        int offset = 0;
        while (matcher.find()) {
            
            if (matchCount == totalMatchCount)
                continue;
            
            int matchEnd = matcher.end() - offset;
            
            String fragment = stringBuilder.toString().substring(0, matchEnd);
            insertRun(
                    firstRun,
                    firstIndex + (matchCount - 1),
                    fragment
            );
            stringBuilder.delete(0, matchEnd);
            offset += matcher.end();
            firstRun.setText(stringBuilder.toString(), 0);
            
            matchCount++;
            
        }
        
    }
    
    private int getMatchCountAndReset(Matcher matcher) {
        
        int getMatchCount = 0; 
        while (matcher.find())
            getMatchCount++;
                
        matcher.reset();
        
        return getMatchCount;
        
    }
    
    private void splitBound(int index, Pattern pattern) {
        
        XWPFRun run = paragraph.getRuns().get(index);
        Matcher matcher = pattern.matcher(run.getText(0));
        if (matcher.find()) {
            insertRun(run, index, matcher.group(1));
            run.setText(matcher.group(2), 0);
            lastIndex++;
        }
        
    }
            
}
