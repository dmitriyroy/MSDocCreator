import javafx.scene.text.FontWeight;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.math.BigInteger;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.xwpf.usermodel.VerticalAlign.BASELINE;

public class Main_1 {

    public static void main(String[] args) {
        String filePathDocTemplateX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.docx";
        String filePathDocX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.docx";
        String filePathXlsX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.xlsx";
//        String filePathDocTemplate ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.doc";
//        String filePathDoc = "D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.doc";


        Order order = new Order();
        order.setCountDay(3);
        order.setOrderId(777);

//        try {
//            createDocX(filePathDocTemplateX, filePathDocX, createDataMap(), order, createProductsInOrder());
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
        List<ProductInOrder> productsInOrder = new ArrayList<>();
        Customer customer = new Customer();
        try {
            createXlsX(filePathXlsX,order,productsInOrder,customer);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static File createDocX(String templateFilePath, String resultFilePath, Map<String, Object> replaceData, Order order, List<ProductInOrder> productsInOrders) throws IOException {
        File resultFile = new File(resultFilePath);
        FileOutputStream fileOutputStream = new FileOutputStream(resultFile);

        FileInputStream fis = new FileInputStream(templateFilePath);
        XWPFDocument templateFile = new XWPFDocument(fis);
        List<XWPFParagraph> paragraphs = templateFile.getParagraphs();



        XWPFDocument document = new XWPFDocument();
        int i = 0;
        for(XWPFParagraph paragraph1 : paragraphs) {
            i++;
            String text = paragraph1.getText();
            if(!text.contains("DATA_TABLE")) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                paragraph.setAlignment(paragraph1.getAlignment());
//                System.out.println("text = " + text);
//                System.out.println("paragraph1.getAlignment() = " + paragraph1.getAlignment());
                paragraph.setFontAlignment(paragraph1.getFontAlignment());
//                System.out.println("paragraph1.getFontAlignment() = " + paragraph1.getFontAlignment());
                paragraph.setNumID(paragraph1.getNumID());
                paragraph.setSpacingAfter(paragraph1.getSpacingAfter());
//                System.out.println("paragraph1.getSpacingAfter() = " + paragraph1.getSpacingAfter());
                paragraph.setSpacingBefore(paragraph1.getSpacingBefore());
//                System.out.println("paragraph1.getSpacingBefore() = " + paragraph1.getSpacingBefore());
                paragraph.setVerticalAlignment(paragraph1.getVerticalAlignment());
//                System.out.println("paragraph1.getVerticalAlignment() = " + paragraph1.getVerticalAlignment());
                run.setText(replaceDataValue(text, replaceData));
//            paragraph.setStyle(paragraph1.getStyle());
//            paragraph.setSpacingAfterLines(paragraph1.getSpacingAfterLines());
//            paragraph.setSpacingBeforeLines(paragraph1.getSpacingBeforeLines());
//            paragraph.setFirstLineIndent(paragraph1.getFirstLineIndent());
//            paragraph.setSpacingLineRule(paragraph1.getSpacingLineRule());
//            paragraph.setSpacingBetween(paragraph1.getSpacingBetween());
            }else{

                createTable(document, order, productsInOrders);
            }
        }

        document.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();
        return resultFile;
    }

    /*
    По принцыпу сметы. Список оборудки сумма, адрес доставки, имя заказчика.
     */
    public static File createXlsX(String fileParth, Order order, List<ProductInOrder> productsInOrders, Customer customer) throws IOException {
        File resultFile = new File(fileParth);
        XSSFWorkbook document = new XSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream(resultFile);

        XSSFSheet sheet = document.createSheet("order_info");
        Row rowOrderNum = sheet.createRow(2);
        Row rowCustomer = sheet.createRow(3);
        Row rowAddress= sheet.createRow(4);
        Font fontBold = document.createFont();
        fontBold.setBold(true);

        XSSFCell cell11 = (XSSFCell) rowOrderNum.createCell(1);
        XSSFCell cell12 = (XSSFCell) rowOrderNum.createCell(2);
        XSSFCell cell13 = (XSSFCell) rowOrderNum.createCell(3);
        XSSFCell cell14 = (XSSFCell) rowOrderNum.createCell(4);

        CellStyle cellStyleLeftBold = document.createCellStyle();
        cellStyleLeftBold.setAlignment(HorizontalAlignment.LEFT);
        cellStyleLeftBold.setFont(fontBold);
        cell12.setCellStyle(cellStyleLeftBold);

        CellStyle cellStyleCenter = document.createCellStyle();
        cellStyleCenter.setAlignment(HorizontalAlignment.CENTER);
        cell13.setCellStyle(cellStyleCenter);
        cell14.setCellStyle(cellStyleCenter);

        cell11.setCellValue("Заказ №");
        cell12.setCellValue(order.getOrderId());
        cell13.setCellValue("от");
        cell12.setCellValue(order.getDateIncomeStringDDMMYYYY());

        document.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();

        return resultFile;
    }

    private static String replaceDataValue(String instring, Map<String, Object> data){
        return instring.replace("ORDER_NUM", data.get("ORDER_NUM").toString())
                    .replace("FIO", data.get("FIO").toString())
                    .replace("ORDER_SUM", data.get("ORDER_SUM").toString())
                    .replace("DAY1", data.get("DAY1").toString())
                    .replace("DAY2", data.get("DAY2").toString())
                    .replace("DAY3", data.get("DAY3").toString())
                    .replace("MONTH1", data.get("MONTH1").toString())
                    .replace("MONTH2", data.get("MONTH2").toString())
                    .replace("MONTH3", data.get("MONTH3").toString())
                    .replace("YEAR1", data.get("YEAR1").toString())
                    .replace("YEAR2", data.get("YEAR2").toString())
                    .replace("YEAR3", data.get("YEAR3").toString())
                    .replace("SNAME", data.get("SNAME").toString())
                    .replace("FNAME", data.get("FNAME").toString())
                    .replace("PNAME", data.get("PNAME").toString());
    }

    private static void createTable(XWPFDocument document, Order order, List<ProductInOrder> productsInOrder){
        XWPFParagraph p0 = document.createParagraph();
        //create table
        XWPFTable table = document.createTable(productsInOrder.size()+1,4);
        table.setCellMargins(100,100,0,100);
        table.setWidth(10000);
        //create first row

        table.getRow(0).getCell(0).removeParagraph(0);
        XWPFParagraph paragraph1 = table.getRow(0).getCell(0).addParagraph();
        paragraph1.setSpacingBefore(0);
        paragraph1.setAlignment(ParagraphAlignment.CENTER);
        paragraph1.setVerticalAlignment(TextAlignment.BOTTOM);
        paragraph1.setSpacingBefore(0);
//        paragraph1.setSpacingBeforeLines(0);
        paragraph1.setSpacingAfter(0);
        XWPFRun run1 = paragraph1.createRun();
        run1.setText("Назва");
        run1.setBold(true);

//        tableRowOne.addNewTableCell();
        table.getRow(0).getCell(1).removeParagraph(0);
        XWPFParagraph paragraph2 = table.getRow(0).getCell(1).addParagraph();
        XWPFRun run2 = paragraph2.createRun();
        paragraph2.setAlignment(ParagraphAlignment.CENTER);
//        paragraph2.setSpacingBefore(0);
        paragraph2.setSpacingAfter(0);
        run2.setText("Ціна за день");
        run2.setBold(true);

//        tableRowOne.addNewTableCell();
        table.getRow(0).getCell(2).removeParagraph(0);
        XWPFParagraph paragraph3 = table.getRow(0).getCell(2).addParagraph();
        paragraph3.setAlignment(ParagraphAlignment.CENTER);
        paragraph3.setSpacingBefore(0);
        paragraph3.setSpacingAfter(0);
        XWPFRun run3 = paragraph3.createRun();
        run3.setText("Кількість днів");
        run3.setBold(true);

//        tableRowOne.addNewTableCell();
        table.getRow(0).getCell(3).removeParagraph(0);
        XWPFParagraph paragraph4 = table.getRow(0).getCell(3).addParagraph();
        XWPFRun run4= paragraph4.createRun();
        paragraph4.setAlignment(ParagraphAlignment.CENTER);
        paragraph4.setSpacingBefore(0);
        paragraph4.setSpacingAfter(0);
        run4.setText("Сума");
        run4.setBold(true);

        table.getRow(0).getCell(0).setColor("CCCCCC");
        table.getRow(0).getCell(1).setColor("CCCCCC");
        table.getRow(0).getCell(2).setColor("CCCCCC");
        table.getRow(0).getCell(3).setColor("CCCCCC");
        table.getRow(0).getCell(0).setWidth("7000");
        table.getRow(0).getCell(1).setWidth("1000");
        table.getRow(0).getCell(2).setWidth("1000");
        table.getRow(0).getCell(3).setWidth("1000");

        //create second row
        for (int i = 1; i <= productsInOrder.size(); i++) {
            table.getRow(i).getCell(0).setText(productsInOrder.get(i-1).getProductName() + " - (" + productsInOrder.get(i-1).getProductModel() + ")");
            table.getRow(i).getCell(1).setText(productsInOrder.get(i-1).getProductPriceString());
            table.getRow(i).getCell(2).setText(order.getCountDay()+"");
            table.getRow(i).getCell(3).setText((productsInOrder.get(i-1).getProductPrice() * order.getCountDay())+"");
        }
    }

    private static Map<String, Object> createDataMap(){
        Map<String,Object> data = new HashMap<>();
        data.put("ORDER_NUM", 777);
        data.put("FIO", "Рой Дмитрий Николаевич");
        data.put("ORDER_SUM", 156.35);
        data.put("DAY1", 25);
        data.put("DAY2", 27);
        data.put("DAY3", 29);
        data.put("MONTH1", 12);
        data.put("MONTH2", 12);
        data.put("MONTH3", 12);
        data.put("YEAR1", 2018);
        data.put("YEAR2", 2018);
        data.put("YEAR3", 2018);
        data.put("SNAME", "Рой");
        data.put("FNAME", "Дмитрий");
        data.put("PNAME", "Николаевич");
        return data;
    }

    private static List<ProductInOrder> createProductsInOrder(){
        List<ProductInOrder> productsInOrder = new ArrayList<>();
        ProductInOrder productInOrder = new ProductInOrder();
        productInOrder.setOrderId(777);
        productInOrder.setProductId(2);
        productInOrder.setProductName("Стул latina");
        productInOrder.setProductModel("Черный");
        productInOrder.setProductPrice(100.00);
        productInOrder.setProductAmount(4);

        productsInOrder.add(productInOrder);

        ProductInOrder productInOrder4 = new ProductInOrder();
        productInOrder4.setOrderId(777);
        productInOrder4.setProductId(4);
        productInOrder4.setProductName("Диван Магнат");
        productInOrder4.setProductModel("Белый");
        productInOrder4.setProductPrice(700.00);
        productInOrder4.setProductAmount(1);

        productsInOrder.add(productInOrder4);

        ProductInOrder productInOrder1 = new ProductInOrder();
        productInOrder1.setOrderId(777);
        productInOrder1.setProductId(3);
        productInOrder1.setProductName("Стол барный");
        productInOrder1.setProductModel("Круглый");
        productInOrder1.setProductPrice(150.00);
        productInOrder1.setProductAmount(2);

        productsInOrder.add(productInOrder1);

        ProductInOrder productInOrder2 = new ProductInOrder();
        productInOrder2.setOrderId(777);
        productInOrder2.setProductId(6);
        productInOrder2.setProductName("Пепельница напольная");
        productInOrder2.setProductModel("");
        productInOrder2.setProductPrice(100.00);
        productInOrder2.setProductAmount(2);

        productsInOrder.add(productInOrder2);

        ProductInOrder productInOrder3 = new ProductInOrder();
        productInOrder3.setOrderId(777);
        productInOrder3.setProductId(70);
        productInOrder3.setProductName("Термопот 20л(Кипятильник)");
        productInOrder3.setProductModel("Бойлер");
        productInOrder3.setProductPrice(250.00);
        productInOrder3.setProductAmount(1);

        productsInOrder.add(productInOrder3);

        return productsInOrder;
    }
}
