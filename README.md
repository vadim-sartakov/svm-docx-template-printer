## Docx template printer.
A set of useful features for printing docx templates. 

### Usage
To fill template with provided POJO simply do:

    Printer<Item> printer = new Printer<>(item, new FileInputStream("template.docx"));
    printer.print(new FileOutputStream("output.docx"));


### Text normalization
By default, when you edit and save text in MSWord it randomly splits text into small fragments,
so called "Runs". Apache POI allows text replacing only on run level. So it would be impossible to make any text replaces
without text normalization. XWPFRunNormalizer class intended to solve this problem. 

     XWPFDocument document = new XWPFDocument(
     new FileInputStream("template.docx"));       
     paragraph = document.getParagraphs().get(0);
     sourceParagraphText = paragraph.getText();
     XWPFRunNormalizer runNormalizer = new XWPFRunNormalizer(paragraph, "\\(.+\\)")
     runNormalizer.normalize();        

### Parameter syntax
It's pretty straightforward to define template parameters. For example 
`[{width: 100} ${manufacturer} ${serialNumber} [{date: "dd.MM.YYYY"} ${releaseDate}]]`

Parameters could be nested. Special attribute `width` will append spaces around result string to match provided width.
Other attributes such as `date` and `number` support default formatters and also can be used in templates.

These parameters are totally eligible:
- `[{width: 20; number: "0.00"} ${price}]`
- `[{date: "dd.MM.YYYY"} ${releaseDate}]`

### Tables
Table rendering also supported. Just define table and insert related collection-like parameter of provided POJO.
There is special parameter `rowNumber` also available.

| №           |	Date                                    | Price                               | Quantity            |
|------------- |----------------------------------------|-------------------------------------|---------------------|
| ${rowNumber} | [{date: "dd.MM.yyyy"} ${history.date}] | [{number: "0.00"} ${history.price}] | ${history.quantity} |
