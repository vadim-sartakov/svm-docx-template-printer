package svm.msoffice.docx.printer.utils;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.List;

/**
 * Dumb class for filling templates.
 * @author sartakov
 */
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

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getDescription() {
        return description;
    }

    public void setDescription(String description) {
        this.description = description;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public String getManufacturer() {
        return manufacturer;
    }

    public void setManufacturer(String manufacturer) {
        this.manufacturer = manufacturer;
    }

    public LocalDate getReleaseDate() {
        return releaseDate;
    }

    public void setReleaseDate(LocalDate releaseDate) {
        this.releaseDate = releaseDate;
    }

    public String getWeight() {
        return weight;
    }

    public void setWeight(String weight) {
        this.weight = weight;
    }

    public String getWidth() {
        return width;
    }

    public void setWidth(String width) {
        this.width = width;
    }

    public String getHeight() {
        return height;
    }

    public void setHeight(String height) {
        this.height = height;
    }

    public String getSerialNumber() {
        return serialNumber;
    }

    public void setSerialNumber(String serialNumber) {
        this.serialNumber = serialNumber;
    }

    public List<History> getHistory() {
        return history;
    }

    public void setHistory(List<History> history) {
        this.history = history;
    }

    public static class History {
        
        public LocalDate date;
        public BigDecimal price;
        public Integer quantity;

        public History(LocalDate date, BigDecimal price, Integer quantity) {
            this.date = date;
            this.price = price;
            this.quantity = quantity;
        }
        
    }
    
}
