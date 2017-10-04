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
import svm.msoffice.docx.printer.impl.Utils;
import svm.msoffice.docx.printer.impl.XWPFRunNormalizer;

/**
 * MS Word template printer.
 * @author sartakov
 * @param <T> 
 */
public class Printer<T> {
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    
    private final T object;
    private final Map<String, Converter> converters;
    private final Map<String, Object> variables = new HashMap<>();
    private final XWPFDocument templateFile;    
    
    public final static Pattern TABLE_CELL_PATTERN = Pattern.compile("\\{([^$\\.:\\}]+)\\.");
    
    private XWPFParagraph paragraph;
    private Map<Integer, Template> templates;
    private int index;
    private int tableRowIndex;
    
    private XWPFTable table;
    private XWPFTableRow row;
    private String tableName;
    private int templateRowIndex;
    private List<Object> list;
        
    public Printer(T object,
            InputStream inputStream,
            Map<String, Converter> converters) {
        
        this.object = object;
        try {
            this.templateFile = new XWPFDocument(new BufferedInputStream(inputStream));
        } catch (IOException e) {
            LOGGER.error("Failed to open file", e);
            throw new RuntimeException(e);
        }
        
        this.converters = converters;
        
    }
        
    public Printer (T object, InputStream inputStream) {
        this(object, inputStream, null);
    }
            
    public void print(OutputStream outputStream) {
        normalizeParameters(templateFile.getParagraphs());
        renderInlineTemplates();
        renderTables();
        save(outputStream);
    }
    
    private void normalizeParameters(List<XWPFParagraph> paragraphs) {
        paragraphs.forEach(currentParagraph -> {
            new XWPFRunNormalizer(currentParagraph, "\\$\\{[^\\{]+\\}").normalize();
            new XWPFRunNormalizer(currentParagraph, "\\[\\{[^\\[\\]]+(?R)\\]").normalize();
        });
    }
        
    private void renderInlineTemplates() {
                
        for (XWPFParagraph currentParagraph : templateFile.getParagraphs()) {
            this.paragraph = currentParagraph;            
            renderTemplatesOfParagraph();
        } 
        
    }
        
    private void renderTemplatesOfParagraph() {
        this.templates = new Parser(this, paragraph).parse();
        renderTemplates();        
    }
        
    private void renderTemplates() {
        templates.entrySet().forEach(entry -> {
            XWPFRun run = paragraph.getRuns().get(entry.getKey());
            run.setText(entry.getValue().render(), 0);
        });     
    }
                        
    private void renderTables() {
        
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
    private void getNameAndValue() {
        
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
    
    private void populateRows() {
                        
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
            
            Utils.copyParagraph(sourceCell.getParagraphArray(0), destinationCell.getParagraphs().get(0));
            
            cellIndex++;
            
        }
        
    }
                
    private void renderRows() {
                     
        tableRowIndex = 0;
        for (XWPFTableRow currentRow : table.getRows()) {
                        
            variables.put("rowNumber", tableRowIndex + 1);
            
            // Bypassing headers
            if (!TABLE_CELL_PATTERN.matcher(currentRow.getCell(0).getText()).find())
                continue;
            
            row = currentRow;
            renderCells();
            
            tableRowIndex++;
            
        }
        
    }
    
    private void renderCells() {
                
        for (XWPFTableCell cell : row.getTableCells()) {
            
            normalizeParameters(cell.getParagraphs());
            for (XWPFParagraph currentParagraph : cell.getParagraphs()) {
                
                paragraph = currentParagraph;
                
                insertIndexInParameters();
                renderTemplatesOfParagraph();
                
            }
        }
        
    }
    
    private void insertIndexInParameters() {
                
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