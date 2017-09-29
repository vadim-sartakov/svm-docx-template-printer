/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.LinkedList;
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
public class XWPFRunNormalizer {
    
    private final XWPFParagraph paragraph;
    private final List<ParamLocation> paramList;
    private final String startParameterBound;
    private final String endParameterBound;
    private final String startBoundFirstSymbol;
    private final Pattern startPattern;
    private final Pattern endPattern;
    /** Maintaining buffer in case parameter start has been splitted in different runs **/
    private final LinkedList<String> runBuffer = new LinkedList<>();
    
    private ParamLocation paramLocation;
    private StringBuilder bufferResultBuilder;
    private String bufferResult = "";
    private String currentText;

    private int index;
    private int startBoundFirstOccurrence = -1;
    private int endBoundFirstOccurrence = -1;
    /** For nested cases */
    private int boundOccurrences = 0;
    private int firstOccurrence;
    private int offset;

    public XWPFRunNormalizer(XWPFParagraph paragraph, String startParameterBound, String endParameterBound) {
        this.paragraph = paragraph;
        this.startParameterBound = startParameterBound;
        this.endParameterBound = endParameterBound;
        this.startBoundFirstSymbol = startParameterBound.substring(0, 1);
        String endBoundLastSymbol = endParameterBound.substring(
                endParameterBound.length() - 1,
                endParameterBound.length());
        this.startPattern = Pattern.compile("(.+)" + "(.+" + Pattern.quote(startBoundFirstSymbol) + ")");
        this.endPattern = Pattern.compile("(.+" + Pattern.quote(endBoundLastSymbol) + ")" + "(.+)");
        this.paramList = new LinkedList<>();
        paramLocation = new ParamLocation();        
    }

    public void normalizeRuns() {
        collectParams();
        replaceRuns();
    }

    class ParamLocation {    
        int start = -1;
        int end = -1;
        StringBuilder paramBuilder = new StringBuilder();
    }

    private void collectParams() {
        
        // Can't change list while iterating. So, accessing elements with index.
        List<XWPFRun> allRuns = paragraph.getRuns();
        for (index = 0; index < allRuns.size(); index++) {

            XWPFRun run = allRuns.get(index);
            currentText = run.getText(0);
            
            detectStartBoundOccurence();
            maintainBuffer();
            checkStart();
            checkEnd();
            appendAndComplete();

        }

    }
    
    private void detectStartBoundOccurence() {
        if (currentText.contains(startBoundFirstSymbol) &&
                boundOccurrences == 0)
            startBoundFirstOccurrence = index;      
    }
    
    private void maintainBuffer() {
        
        runBuffer.offer(currentText);
        if (runBuffer.size() > startParameterBound.length())
            runBuffer.poll();
        
        bufferResultBuilder = new StringBuilder();
        runBuffer.forEach(string -> bufferResultBuilder.append(string));
        bufferResult = bufferResultBuilder.toString();

    }

    private void checkStart() {

        if (bufferResult.contains(startParameterBound)) {

            if (paramLocation.start == -1) {
                startBoundFirstOccurrence =
                        splitBoundAndRest(startBoundFirstOccurrence, 1);                
                paramLocation.start = startBoundFirstOccurrence;
                currentText = startParameterBound;
            }

            runBuffer.clear();
            boundOccurrences++;

        }

    }
        
    private int splitBoundAndRest(int firstOccurrence, int offset) {

        boolean isStartBound = offset == 1;
        XWPFRun firstOccurrenceRun = paragraph.getRuns().get(firstOccurrence);
        String firstOccurrenceRunText = firstOccurrenceRun.getText(0);

        Pattern pattern = isStartBound ? startPattern : endPattern;
        Matcher matcher = pattern.matcher(firstOccurrenceRunText);
        
        if (!matcher.find())
            return firstOccurrence;
        
        String boundTrace, rest;
        if (isStartBound) {
            boundTrace = matcher.group(2);
            rest = matcher.group(1);
        } else {
            boundTrace = matcher.group(1);
            rest = matcher.group(2);
        }        
                
        this.firstOccurrence = firstOccurrence;
        this.offset = offset;
        
	firstOccurrenceRun.setText(rest, 0);
	addBoundRun(boundTrace); 
        
        return this.firstOccurrence;

    }

    private void addBoundRun(String boundTrace) {
        
        XWPFRun firstRun = paragraph.getRuns().get(firstOccurrence);
        firstOccurrence += offset;
        index += offset;
        XWPFRun newRun = paragraph.insertNewRun(firstOccurrence);
        Utils.copyRun(firstRun, newRun);
        newRun.setText(boundTrace);
        
    }
    
    private void checkEnd() {
        
        detectEndBoundOccurence();
        if (paramLocation.start != -1 && currentText.contains(endParameterBound)) {
                        
            if (boundOccurrences == 1) {
                splitBoundAndRest(endBoundFirstOccurrence, 0);
                paramLocation.end = index;
                currentText = endParameterBound;
            } else
                boundOccurrences--;
            
        }
        
    }
    
    private void detectEndBoundOccurence() {         
        if (currentText.contains(endParameterBound) &&
                boundOccurrences == 1)
            endBoundFirstOccurrence = index;
    }

    private void appendAndComplete() {
        if (paramLocation.start >= 0)
            paramLocation.paramBuilder.append(currentText);       

        if (paramLocation.end >= 0) {
            paramList.add(paramLocation);
            paramLocation = new ParamLocation();
            startBoundFirstOccurrence = -1;
            endBoundFirstOccurrence = -1;
            boundOccurrences = 0;
        }
    }

    private void replaceRuns() {

        offset = 0;
        for (ParamLocation currentParam : paramList) {

            // Taking in count offset after deletion
            int start = currentParam.start - offset;
            int end = currentParam.end - offset;

            XWPFRun currentRun = paragraph.getRuns().get(start);
            XWPFRun newRun = paragraph.insertNewRun(end + 1);

            Utils.copyRun(currentRun, newRun);
            newRun.setText(currentParam.paramBuilder.toString(), 0);

            for (int i = start; i <= end; i++) {
                paragraph.removeRun(start);
                if (i < end)
                    offset++;
            }

        }

    }

}
