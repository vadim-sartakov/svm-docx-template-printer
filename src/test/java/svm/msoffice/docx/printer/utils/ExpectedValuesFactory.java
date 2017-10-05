package svm.msoffice.docx.printer.utils;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.AbstractMap;
import java.util.HashMap;
import java.util.Map;
import svm.msoffice.docx.printer.impl.Template;

/**
 *
 * @author sartakov
 */
public class ExpectedValuesFactory {
    
    public static Map<Integer, Template> getCorrectTemplates() {
        
        Map<Integer, Template> correctTemplates = new HashMap<>();
        Template template, enclosingTemplate;
        
        template = new Template("${name}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${name}", "Screwdriver");
        correctTemplates.put(2, template);
        
        template = new Template("[{width: 20; number: \"0.00\"} ${price}]");
        template.format = new AbstractMap.SimpleEntry<>("number", "0.00");
        template.width = 20;
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${price}", new BigDecimal("12.6586"));
        correctTemplates.put(4, template);
                
        template = new Template("${description}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${description}", "Handy screwdriver");
        correctTemplates.put(6, template);
        
        enclosingTemplate = new Template("[{date: \"dd.MM.YYYY\"} ${releaseDate}]");
        enclosingTemplate.format = new AbstractMap.SimpleEntry<>("date", "dd.MM.YYYY");
        enclosingTemplate.parameterValues = new HashMap<>();
        enclosingTemplate.parameterValues.put("${releaseDate}", LocalDate.of(2012, 3, 25));
        
        template = new Template("[{width: 100} ${manufacturer} ${serialNumber} [{date: \"dd.MM.YYYY\"} ${releaseDate}]]");
        template.width = 100;
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${manufacturer}", "Some factory lmtd");
        template.parameterValues.put("${serialNumber}", "15358-548");
        template.enclosingTemplate = enclosingTemplate;
        correctTemplates.put(8, template);
        
        template = new Template("${weight}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${weight}", "0.5 kg");
        correctTemplates.put(10, template);
        
        template = new Template("${height}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${height}", "20 mm");
        correctTemplates.put(12, template);
        
        template = new Template("${width}");
        template.parameterValues = new HashMap<>();
        template.parameterValues.put("${width}", "500 mm");
        correctTemplates.put(14, template);
              
        return correctTemplates;
        
    }
    
    public static Map<Integer, String> getRenderResults() {
        
        Map<Integer, String> result = new HashMap<>();
        result.put(2, "Screwdriver");
        result.put(4, "       12,66        ");
        result.put(6, "Handy screwdriver");
        result.put(8, "                               Some factory lmtd 15358-548 25.03.2012                               ");
        result.put(10, "0.5 kg");
        result.put(12, "20 mm");
        result.put(14, "500 mm");
        
        return result;
        
    }
    
    public static Item getItem() {

        Item item = new Item();
        item.setName("Screwdriver");
        item.setDescription("Handy screwdriver");
        item.setPrice(new BigDecimal("12.6586"));
        item.setManufacturer("Some factory lmtd");
        item.setSerialNumber("15358-548");
        item.setReleaseDate(LocalDate.of(2012, 3, 25));
        item.setWeight("0.5 kg");
        item.setHeight("20 mm");
        item.setWidth("500 mm");

        return item;
        
    }
    
}
