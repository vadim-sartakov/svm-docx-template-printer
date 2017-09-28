/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer.impl;

import java.util.LinkedList;
import java.util.List;
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
    private final List<ParamMeta> paramList;
    private ParamMeta param;
    private final String startParameterBound;
    private final String endParameterBound;

    // Maintaining buffer in case parameter start has been splitted in different runs
    private final LinkedList<String> runBuffer = new LinkedList<>();
    private StringBuilder bufferResultBuilder;
    private String bufferResult;
    private String currentText;
    private String textToAppend;

    private int index;
    private int startBoundFirstOccurrence;
    // For nested cases
    private int boundOccurrences;

    public XWPFRunNormalizer(XWPFParagraph paragraph, String startParameterBound, String endParameterBound) {
        this.paragraph = paragraph;
        this.startParameterBound = startParameterBound;
        this.endParameterBound = endParameterBound;
        this.paramList = new LinkedList<>();
        param = new ParamMeta();        
    }

    public void normalizeRuns() {
        collectParams();
        replaceRuns();
    }

    class ParamMeta {    
        int start = -1;
        int end = -1;
        StringBuilder paramBuilder = new StringBuilder();
    }

    private void collectParams() {

        index = -1;
        startBoundFirstOccurrence = -1;
        boundOccurrences = 0;

        for (XWPFRun currentRun : paragraph.getRuns()) {

            index++;
            currentText = currentRun.getText(0);
            textToAppend = currentText;

            if (currentText.contains(startParameterBound.substring(0, 1))
                    && boundOccurrences == 0)
                startBoundFirstOccurrence = index;

            maintainBuffer();
            checkStart();
            checkEnd();
            appendAndComplete();

        }

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

            runBuffer.clear();

            if (param.start == -1)
                param.start = startBoundFirstOccurrence;

            if (bufferResult.length() == startParameterBound.length())
                textToAppend = bufferResult;

            boundOccurrences++;

        }

    }

    private void checkEnd() {
        if (param.start != -1 && currentText.contains(endParameterBound))
            if (boundOccurrences == 1)
                param.end = index;
            else
                boundOccurrences--;
    }

    private void appendAndComplete() {
        if (param.start >= 0)
            param.paramBuilder.append(textToAppend);       

        if (param.end >= 0) {
            paramList.add(param);
            param = new ParamMeta();
            startBoundFirstOccurrence = -1;
            boundOccurrences = 0;
        }
    }

    private void replaceRuns() {

        int offset = 0;
        for (ParamMeta currentParam : paramList) {

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
