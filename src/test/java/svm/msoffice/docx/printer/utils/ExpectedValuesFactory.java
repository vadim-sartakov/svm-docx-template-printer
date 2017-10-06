package svm.msoffice.docx.printer.utils;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.AbstractMap;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Map;
import svm.msoffice.docx.printer.impl.Table;
import svm.msoffice.docx.printer.impl.Template;

/**
 *
 * @author sartakov
 */
public class ExpectedValuesFactory {
    
    public static Map<Integer, Template> getCorrectTemplates() {
        
        Map<Integer, Template> correctTemplates = new HashMap<>();
        Template template, enclosingTemplate;
        
        template = new Template("${name}", "Screwdriver");
        correctTemplates.put(2, template);
        
        template = new Template("[{width: 20; number: \"0.00\"} ${price}]");
        template.format = new AbstractMap.SimpleEntry<>("number", "0.00");
        template.width = 20;
        template.parameterValues.put("${price}", new BigDecimal("12.6586"));
        correctTemplates.put(4, template);
                
        template = new Template("${description}", "Handy screwdriver");
        correctTemplates.put(6, template);
        
        enclosingTemplate = new Template("[{date: \"dd.MM.YYYY\"} ${releaseDate}]");
        enclosingTemplate.format = new AbstractMap.SimpleEntry<>("date", "dd.MM.YYYY");
        enclosingTemplate.parameterValues.put("${releaseDate}", LocalDate.of(2012, 3, 25));
        
        template = new Template("[{width: 100} ${manufacturer} ${serialNumber} [{date: \"dd.MM.YYYY\"} ${releaseDate}]]");
        template.width = 100;
        template.parameterValues.put("${manufacturer}", "Some factory lmtd");
        template.parameterValues.put("${serialNumber}", "15358-548");
        template.enclosingTemplate = enclosingTemplate;
        correctTemplates.put(8, template);
        
        template = new Template("${weight}", "0.5 kg");
        correctTemplates.put(10, template);
        
        template = new Template("${height}", "20 mm");
        correctTemplates.put(12, template);
        
        template = new Template("${width}", "500 mm");
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
        
        item.history = new LinkedList<>();
        item.history.add(new Item.History(
                LocalDate.of(2011, 11, 10),
                new BigDecimal("5.25"),
                10)
        );
        item.history.add(new Item.History(
                LocalDate.of(2012, 2, 5),
                new BigDecimal("6.20"),
                12)
        );
        item.history.add(new Item.History(
                LocalDate.of(2012, 4, 26),
                new BigDecimal("6.80"),
                22)
        );
        item.history.add(new Item.History(
                LocalDate.of(2015, 12, 31),
                new BigDecimal("7.5"),
                35)
        );

        return item;
        
    }
    
    public static Table getCorrectTable() {
        
        Item item = getItem();
        
        Table table = new Table();
        int index = 1;
        for (Item.History historyItem : item.getHistory()) {
            
            Table.Row newRow = table.addRow(index);
            Template template;
            
            template = new Template("${rowNumber}", index);
            newRow.addCell(0, template);
            
            template = new Template("${history[" + (index - 1) + "].date}");
            template.format = new AbstractMap.SimpleEntry<>("date", "dd.MM.yyyy");
            template.parameterValues.put("${history[" + (index - 1) + "].date}", historyItem.getDate());
            
            newRow.addCell(1, template);
            
            template = new Template("${history[" + (index - 1) + "].price}");
            template.format = new AbstractMap.SimpleEntry<>("number", "0.00");
            template.parameterValues.put("${history[" + (index - 1) + "].price}", historyItem.getPrice());
            
            newRow.addCell(2, template);
            newRow.addCell(3, new Template("${history[" + (index - 1) + "].quantity}", historyItem.getQuantity()));

            index++;
            
        }
        
        return table;
        
    }
        
}