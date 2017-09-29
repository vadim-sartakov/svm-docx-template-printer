/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.LinkedList;
import java.util.List;
import java.util.ListIterator;
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
    private ParamLocation param;
    private final String startParameterBound;
    private final String endParameterBound;
    private final String startBoundFirstSymbol;

    /** Maintaining buffer in case parameter start has been splitted in different runs **/
    private final LinkedList<String> runBuffer = new LinkedList<>();
    private StringBuilder bufferResultBuilder;
    private String bufferResult = "";
    private String currentText;

    private int index;
    private int startBoundFirstOccurrence = -1;
    private int endBoundFirstOccurrence = -1;
    /** For nested cases */
    private int boundOccurrences = 0;
    
    private String bound;
    private int firstOccurrence;
    private int offset;

    public XWPFRunNormalizer(XWPFParagraph paragraph, String startParameterBound, String endParameterBound) {
        this.paragraph = paragraph;
        this.startParameterBound = startParameterBound;
        this.endParameterBound = endParameterBound;
        this.startBoundFirstSymbol = startParameterBound.substring(0, 1);
        this.paramList = new LinkedList<>();
        param = new ParamLocation();        
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

            detectBoundOccurences();            
            maintainBuffer();
            checkStart();
            checkEnd();
            appendAndComplete();

        }

    }
    
    private void detectBoundOccurences() {
          
        if (currentText.contains(startBoundFirstSymbol) &&
                boundOccurrences == 0)
            startBoundFirstOccurrence = index;
        
        if (currentText.contains(endParameterBound) &&
                boundOccurrences == 1)
            endBoundFirstOccurrence = index;
        
    }

    private void maintainBuffer() {
            
        runBuffer.offer(currentText);
        if (runBuffer.size() > startParameterBound.length())
            runBuffer.poll();
        
        bufferResultBuilder = new StringBuilder();
        runBuffer.forEach(string -> bufferResultBuilder.append(string));
        bufferResult = bufferResultBuilder.toString();

        if (bufferResult.equals(startParameterBound))
            currentText = bufferResult;

    }

    private void checkStart() {

        if (bufferResult.contains(startParameterBound)) {

            if (param.start == -1) {
                startBoundFirstOccurrence =
                        splitBoundAndRest(startParameterBound, startBoundFirstOccurrence, 1);                
                param.start = startBoundFirstOccurrence;
            }

            runBuffer.clear();
            boundOccurrences++;

        }

    }
    
    private int splitBoundAndRest(String bound, int firstOccurrence, int offset) {

        if (bufferResult.equals(startParameterBound))
            return firstOccurrence;
        
        this.bound = bound;
        this.firstOccurrence = firstOccurrence;
        this.offset = offset;
        
	clearFirstOccurrenceFromBound();
	addBoundRun();
        
        currentText = bound;
        
        return this.firstOccurrence;

    }

    private void clearFirstOccurrenceFromBound() {

        XWPFRun firstOccurrenceRun = paragraph.getRuns().get(firstOccurrence);
        String firstOccurrenceRunText = firstOccurrenceRun.getText(0);

        String runTextWithoutBound = firstOccurrenceRunText;
        for (int i = 0; i < bound.length(); i++)
            runTextWithoutBound = runTextWithoutBound.replace(bound.substring(i, i + 1), "");
        
        firstOccurrenceRun.setText(runTextWithoutBound, 0);

    }

    private void addBoundRun() {
        
        XWPFRun firstRun = paragraph.getRuns().get(firstOccurrence);
        firstOccurrence += offset;
        index += offset;
        XWPFRun newRun = paragraph.insertNewRun(firstOccurrence);
        Utils.copyRun(firstRun, newRun);
        newRun.setText(bound);
        
    }

    private void checkEnd() {
        
        if (param.start != -1 && currentText.contains(endParameterBound)) {
                        
            if (boundOccurrences == 1) {
                splitBoundAndRest(endParameterBound, endBoundFirstOccurrence, 0);
                param.end = index;
            } else
                boundOccurrences--;
            
        }
        
    }

    private void appendAndComplete() {
        if (param.start >= 0)
            param.paramBuilder.append(currentText);       

        if (param.end >= 0) {
            paramList.add(param);
            param = new ParamLocation();
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
