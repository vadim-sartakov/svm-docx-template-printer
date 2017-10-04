/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package svm.msoffice.docx.printer;

/**
 *
 * @author sartakov
 */
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.beanutils.PropertyUtils;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import svm.msoffice.docx.printer.impl.Converter;
import svm.msoffice.docx.printer.impl.Parser;
import svm.msoffice.docx.printer.impl.Template;
import svm.msoffice.docx.printer.impl.XWPFRunNormalizer;

/**
 * MS Word template printer.
 * @author sartakov
 * @param <T> 
 */
public class Printer<T> {
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    
    private T object;
    private XWPFDocument templateFile;
    private OutputStream outputStream;
    
    /**
     * Format and parameters
     */
    public final static Pattern FORMAT_SCOPE_PATTERN = Pattern.compile("\\[(\\{[^$]*\\}) +(.+)\\]");
    /**
     * Parameters only, without format scope
     */
    public final static Pattern PARAMETER_SCOPE_PATTERN = Pattern.compile("(\\$\\{([\\w.\\[\\]]+)\\})(?!.*\\])");
    /**
     * Format - value
     */
    public final static Pattern FORMAT_PATTERN = Pattern.compile("([\\w-]+): *\"?([\\wА-Яа-я .']+)\"?");
    /**
     * Parameter - content
     */
    public final static Pattern PARAMETER_PATTERN = Pattern.compile("\\$\\{([\\w.\\[\\]]+)\\}");
    public final static Pattern TABLE_CELL_PATTERN = Pattern.compile("\\{([^$\\.:\\}]+)\\.");
    
    private XWPFParagraph paragraph;
    private Map<String, Converter> converters;
    private Map<Integer, Template> templates;
    private final Map<String, Object> variables = new HashMap<>();
    private int index;
    private int tableRowIndex;
    
    private XWPFTable table;
    private XWPFTableRow row;
    private String tableName;
    private int templateRowIndex;
    private List<Object> list;
        
    public Printer(T object,
            InputStream inputStream,
            OutputStream outputStream,
            Map<String, Converter> converters) throws IOException {
        
        this.object = object;
        this.templateFile = new XWPFDocument(new BufferedInputStream(inputStream));
        this.outputStream = outputStream;
        this.converters = converters;
        
    }
        
    public Printer (T object, InputStream inputStream, OutputStream outputStream) throws IOException {
        this(object, inputStream, outputStream, null);
    }
            
    public void print() throws Exception {
        
        normalizeParameters(templateFile.getParagraphs());
        parseTemplates();
        renderInlineTemplates();
        renderTables();
        
    }
    
    private void normalizeParameters(List<XWPFParagraph> paragraphs) {
        paragraphs.forEach(currentParagraph -> {
            new XWPFRunNormalizer(currentParagraph, "\\$\\{[^\\{]+\\}").normalize();
            new XWPFRunNormalizer(currentParagraph, "\\[\\{[^\\[\\]]+(?R)\\]").normalize();
        });
    }
    
    private void parseTemplates() {
        templateFile.getParagraphs().forEach(currentParagraph -> {
            new Parser(this, currentParagraph).parse();
        });
    }
    
    private void renderInlineTemplates() throws Exception {
                
        for (XWPFParagraph currentParagraph : templateFile.getParagraphs()) {
            this.paragraph = currentParagraph;            
            renderTemplatesOfParagraph();
        } 
        
    }
        
    private void renderTemplatesOfParagraph() throws Exception {
        
        if (!PARAMETER_PATTERN.matcher(paragraph.getText()).find())
            return;
            
        this.templates = new HashMap<>();

        try {
            parseTemplates();
        } catch (Exception e) {
            LOGGER.error("Failed to parse parameters", e);
            throw e;
        }

        renderTemplates();
        
    }
        
    private void renderTemplates() {
        
        if (templates.isEmpty())
            return;
        
        index = 0;
        for (XWPFRun currentRun : paragraph.getRuns()) {
        
            Template currentTemplate = templates.get(index);
            if (currentTemplate != null)
                currentTemplate.render();
            
            index++;
        
        }
        
    }
                        
    private void renderTables() throws Exception {
        
        for (XWPFTable currentTable : templateFile.getTables()) {
            table = currentTable;
            getNameAndValue();
            populateRows();
            renderRows();
        }
        
    }
    
    /**
     * Any column in first section (before first dot)
     * should contain table name to iterate.
     */
    private void getNameAndValue() throws Exception {
        
        index = 0;
        for (XWPFTableRow currentRow : table.getRows()) {
            
            row = currentRow;
            findTableName();
            if (!tableName.isEmpty()) {
                templateRowIndex = index;
                break;
            }
            
            index++;
            
        }
        
        if (tableName.isEmpty())
            return;
        
        try {
            list = (List) PropertyUtils.getProperty(object, tableName);
        } catch (Exception e) {
            LOGGER.warn("Failed to get list by name {}", tableName);
        }
        
        list = list == null ? new LinkedList<>() : list;
        
    }
    
    private void findTableName() {
                
        for (XWPFTableCell cell : row.getTableCells()) {
            
            String firstCellText = cell.getText();
            Matcher matcher = TABLE_CELL_PATTERN.matcher(firstCellText);
            if (matcher.find()) {
                tableName = matcher.group(1);
                break;
            } else
                tableName = "";
            
        }        
    
    }
    
    private void populateRows() throws Exception {
                        
        for (int i = 0; i < list.size() - 1; i++) {
            XWPFTableRow newRow = table.createRow();
            copyRow(table.getRow(templateRowIndex), newRow);
        }
            
    }
    
    // Applying ctr copying leads to unexpected results. So, copying by hand.
    private void copyRow(XWPFTableRow source, XWPFTableRow destination) {
        
        int cellIndex = 0;
        for (XWPFTableCell sourceCell : source.getTableCells()) {
            
            XWPFTableCell destinationCell = destination.getCell(cellIndex);
            destinationCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            
            copyParagraph(sourceCell.getParagraphArray(0), destinationCell.getParagraphs().get(0));
            
            cellIndex++;
            
        }
        
    }
    
    private void copyParagraph(XWPFParagraph sourceParagraph, XWPFParagraph destinationParagraph) {
        destinationParagraph.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
        sourceParagraph.getRuns().forEach(sourceRun -> {
            copyRun(sourceRun, destinationParagraph.createRun());
        });
    }
    
    private void copyRun(XWPFRun sourceRun, XWPFRun destinationRun) {
        destinationRun.getCTR().setRPr(sourceRun.getCTR().getRPr());
        destinationRun.setText(sourceRun.getText(0), 0);
    }
        
    private void renderRows() throws Exception {
                     
        tableRowIndex = 0;
        for (XWPFTableRow currentRow : table.getRows()) {
                        
            variables.put("rowNumber", tableRowIndex + 1);
            
            // Bypassing headers
            if (!PARAMETER_PATTERN.matcher(currentRow.getCell(0).getText()).find())
                continue;
            
            row = currentRow;
            renderCells();
            
            tableRowIndex++;
            
        }
        
    }
    
    private void renderCells() throws Exception {
                
        for (XWPFTableCell cell : row.getTableCells()) {
            
            normalizeParameters(cell.getParagraphs());
            for (XWPFParagraph currentParagraph : cell.getParagraphs()) {
                
                paragraph = currentParagraph;
                
                insertIndexInParameters();
                renderTemplatesOfParagraph();
                
            }
        }
        
    }
    
    private void insertIndexInParameters() throws Exception {
                
        for (XWPFRun currentRun : paragraph.getRuns()) {
            
            String runText = currentRun.getText(0);
            Matcher matcher = TABLE_CELL_PATTERN.matcher(runText);
            if (!matcher.find())
                continue;

            currentRun.setText(runText.replaceAll("\\{" + tableName,
                    "\\{" + tableName + "[" + tableRowIndex + "]"), 0);
            
        }
        
    }
    
    public void save(OutputStream outputStream) {
        
        try {
            templateFile.write(new BufferedOutputStream(outputStream));
        } catch (IOException e) {
            LOGGER.error("Failed to save output file", e);
        } finally {
            IOUtils.closeQuietly(templateFile);
            IOUtils.closeQuietly(outputStream);
        }
        
    }
    
    public void save(String fileName) {
                
        try {
            save(new FileOutputStream(fileName));
        } catch (IOException e) {
            LOGGER.error("Error creating file " + fileName, e);
        }
        
    }
    
    public T getObject() {
        return object;
    }

    public Map<String, Object> getVariables() {
        return variables == null ? null : Collections.unmodifiableMap(variables);
    }

    public Map<String, Converter> getConverters() {
        return converters == null ? null : Collections.unmodifiableMap(converters);
    }

    public XWPFDocument getTemplateFile() {
        return templateFile;
    }
    
}