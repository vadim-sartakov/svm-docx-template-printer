package svm.msoffice.docx.printer;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import svm.msoffice.docx.printer.impl.DataHolder;
import svm.msoffice.docx.printer.impl.Table;
import svm.msoffice.docx.printer.impl.TableParser;
import svm.msoffice.docx.printer.impl.Template;
import svm.msoffice.docx.printer.impl.TemplateParser;
import svm.msoffice.docx.printer.impl.XWPFRunNormalizer;

/**
 * MS Word template printer.
 * @author sartakov
 * @param <T> 
 */
public class Printer<T> {
    
    private final static Logger LOGGER = LoggerFactory.getLogger(Printer.class);
    
    private final DataHolder dataHolder;
    private final XWPFDocument templateFile;
                
    public Printer(T object,
            InputStream inputStream,
            Map<String, Converter> converters) {
        
        this.dataHolder = new DataHolder(object, converters);

        try {
            this.templateFile = new XWPFDocument(new BufferedInputStream(inputStream));
        } catch (IOException e) {
            LOGGER.error("Failed to open file", e);
            throw new RuntimeException(e);
        }
        
    }
        
    public Printer(T object, InputStream inputStream) {
        this(object, inputStream, null);
    }
            
    public void print(OutputStream outputStream) {
        normalizeParameters(templateFile.getParagraphs());
        parseAndRenderInlineTemplates();
        parseAndRenderTables();
        save(outputStream);
    }
    
    private void normalizeParameters(List<XWPFParagraph> paragraphs) {
        paragraphs.forEach(currentParagraph -> {
            XWPFRunNormalizer.normalizeParameters(currentParagraph);
        });
    }
        
    private void parseAndRenderInlineTemplates() {       
        templateFile.getParagraphs().forEach(paragraph -> {
            Map<Integer, Template> templates = new TemplateParser(dataHolder, paragraph).parse();
            Template.renderTemplates(templates, paragraph);
        });
    }
                        
    private void parseAndRenderTables() {
        templateFile.getTables().forEach(xwpfTable -> {
            Table table = new TableParser(dataHolder, xwpfTable).parse();
            table.render(xwpfTable);
        });
    }
        
    public void save(OutputStream outputStream) {
        
        try {
            templateFile.write(new BufferedOutputStream(outputStream));
        } catch (IOException e) {
            LOGGER.error("Failed to save template to output stream", e);
        } finally {
            IOUtils.closeQuietly(templateFile);
            IOUtils.closeQuietly(outputStream);
        }
        
    }
    
    public void save(String fileName) {
                
        try {
            save(new FileOutputStream(fileName));
        } catch (IOException e) {
            LOGGER.error("Failed to save output file " + fileName, e);
        }
        
    }

    public XWPFDocument getTemplateFile() {
        return templateFile;
    }
        
}