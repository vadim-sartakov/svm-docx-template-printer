package svm.msoffice.docx.printer.impl;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
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
    private final List<XWPFRun> allRuns;
    private final Pattern pattern;   
    private final Map<MatchPair, String> matches = new HashMap<>();
    
    private XWPFRun firstRun, lastRun;
    private int firstIndex, lastIndex;
    
    public XWPFRunNormalizer(XWPFParagraph paragraph, String regex) {
        this(paragraph, regex, ".+");
    }
            
    /**
     * 
     * @param paragraph
     * @param regex
     * @param boundSplitRestRegex - regex for capturing rest in the bound splitting procedure.
     */
    public XWPFRunNormalizer(XWPFParagraph paragraph,
            String regex,
            String boundSplitRestRegex) {
        this.paragraph = paragraph;
        this.allRuns = paragraph.getRuns();
        this.pattern = Pattern.compile(regex);
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
            int matchCount = 0;
            while (matcher.find()) {
                                
                MatchPair matchPair = new MatchPair(matcher.start(), matcher.end());
                
                if (matches.get(matchPair) != null)
                    continue;
                
                matchCount++;
                if (matchCount > 1)
                    lastIndex++;
                
                String matchedString = levelOneBuffer.substring(matcher.start(), matcher.end());
                matches.put(matchPair, matchedString);
                runBackwardLoop(matchedString);
                                
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
                splitBounds(matchedString);
                
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
    
    private void insertRun(XWPFRun sourceRun, int index, String newRunText) {
        XWPFRun newRun = paragraph.insertNewRun(index);
        Utils.copyRun(sourceRun, newRun);
        newRun.setText(newRunText, 0);
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
            offset += matchEnd;
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
    
    private void splitBounds(String matchedString) {
        
        XWPFRun run = paragraph.getRuns().get(lastIndex);
        String runText = run.getText(0);
        Matcher matcher = Pattern.compile(Pattern.quote(matchedString))
                .matcher(runText);
        matcher.find();
        
        boolean splitted = false;
        if (matcher.start() > 0) {
            insertRun(run, lastIndex, runText.substring(0, matcher.start()));
            splitted = true;
            lastIndex++;
        }
        
        if (matcher.end() < runText.length()) {
            insertRun(run, lastIndex + 1, runText.substring(matcher.end(), runText.length()));
            splitted = true;
            lastIndex++;
        }
        
        if (splitted) 
            run.setText(matchedString, 0);
        
    }
            
}

class MatchPair {
    
    final int start;
    final int end;

    public MatchPair(int start, int end) {
        this.start = start;
        this.end = end;
    }

    @Override
    public int hashCode() {
        int hash = 7;
        hash = 67 * hash + this.start;
        hash = 67 * hash + this.end;
        return hash;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj) {
            return true;
        }
        if (obj == null) {
            return false;
        }
        if (getClass() != obj.getClass()) {
            return false;
        }
        final MatchPair other = (MatchPair) obj;
        if (this.start != other.start) {
            return false;
        }
        if (this.end != other.end) {
            return false;
        }
        return true;
    }
        
}