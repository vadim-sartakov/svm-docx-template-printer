package svm.msoffice.docx.printer.utils;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.List;
import lombok.AllArgsConstructor;
import lombok.Data;

@Data
public class Item {
    
    String name;
    String description;
    BigDecimal price;
    String serialNumber;
    String manufacturer;
    LocalDate releaseDate;
    String weight;
    String width;
    String height;
    List<History> history;

    @Data
    @AllArgsConstructor
    public static class History {
        
        private LocalDate date;
        private BigDecimal price;
        private Integer quantity;

    }
    
}
