import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.awt.Color;

public class Main_1 {

    public static void main(String[] args) {
        String documentFolderPath = "D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents";
        String filePathDocTemplateX = documentFolderPath + "/dogovor_fiz_template.docx";
        String filePathDocX = documentFolderPath + "/Contract_fiz_order-";
        String filePathXlsX = documentFolderPath + "/Estimate_order-";
        String filePathXlsX2 = "/Smeta_EventHelp_order-";
//        String filePathDocTemplate ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.doc";
//        String filePathDoc = "D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.doc";

        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

        Order order = new Order();
        order.setCountDay(3);
        order.setOrderId(777);
        order.setDeliveryAddress("г.Днепр, ул.Титова, д.14А, кв.55");
        try {
            order.setDateIncome(dateFormat.parse("2019-01-01"));
            order.setDateStart(dateFormat.parse("2019-01-04"));
            order.setDateShipOut(dateFormat.parse("2019-01-03"));
            order.setDateShipBack(dateFormat.parse("2019-01-05"));
        } catch (ParseException e) {
            e.printStackTrace();
        }
        Customer customer = new Customer();
        customer.setfName("Рой");
        customer.setsName("Дмитрий");
        customer.setpName("Николаевич");

        List<ProductInOrder> productsInOrder = createProductsInOrder();

//        createXlsX(filePathXlsX + order.getOrderId() + ".xlsx", order, productsInOrder, customer);
        createXlsX2(documentFolderPath, filePathXlsX2 + order.getOrderId() + ".xlsx", order, productsInOrder, customer);
//        try {
//            createDocX(filePathDocTemplateX, filePathDocX + order.getOrderId() + ".docx", createDataMap(), order, productsInOrder);
//        } catch (IOException e) {
//            e.printStackTrace();
//        }

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
    2 аркуш
     */
    public static File createXlsX2(String documentFolderPath, String outFileName, Order order, List<ProductInOrder> productsInOrders, Customer customer){
        File resultFile = new File(documentFolderPath + outFileName);
        XSSFWorkbook document = new XSSFWorkbook();
        FileOutputStream fileOutputStream = null;

        int rowWidth = 4500;

        try {
            fileOutputStream = new FileOutputStream(resultFile);
            XSSFSheet sheet = document.createSheet("order_info");
            Row rowOrderNum = sheet.createRow(1);
            Row rowProject = sheet.createRow(2);
            Row rowCustomer = sheet.createRow(3);
            Row rowDateStart = sheet.createRow(4);
            Row rowAddress = sheet.createRow(5);
            Row rowContractor = sheet.createRow(6);

            CellRangeAddress cellRangeOrderNum = new CellRangeAddress(1, 1, 4, 7);
            CellRangeAddress cellRangeProject = new CellRangeAddress(2, 2, 4, 7);
            CellRangeAddress cellRangeSuctomer = new CellRangeAddress(3, 3, 4, 7);
            CellRangeAddress cellRangeDateStart = new CellRangeAddress(4, 4, 4, 7);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(5, 5, 4, 7);
            CellRangeAddress cellRangeContractor = new CellRangeAddress(6, 6, 4, 7);
            sheet.addMergedRegion(cellRangeOrderNum);
            sheet.addMergedRegion(cellRangeProject);
            sheet.addMergedRegion(cellRangeSuctomer);
            sheet.addMergedRegion(cellRangeDateStart);
            sheet.addMergedRegion(cellRangeAddress);
            sheet.addMergedRegion(cellRangeContractor);
            sheet.setColumnWidth(0, rowWidth);
            sheet.setColumnWidth(1, rowWidth);
            sheet.setColumnWidth(2, rowWidth);
            sheet.setColumnWidth(3, rowWidth);
            sheet.setColumnWidth(4, rowWidth);
            sheet.setColumnWidth(5, rowWidth);
            sheet.setColumnWidth(6, rowWidth);
            sheet.setColumnWidth(7, rowWidth);
            sheet.setColumnWidth(8, rowWidth);
            Font fontBold = document.createFont();
            fontBold.setBold(true);

            XSSFCell cell11 = (XSSFCell) rowOrderNum.createCell(3);
            XSSFCell cell12 = (XSSFCell) rowOrderNum.createCell(4);
            XSSFCell cell13 = (XSSFCell) rowOrderNum.createCell(5);
            XSSFCell cell14 = (XSSFCell) rowOrderNum.createCell(6);
            XSSFCell cell15 = (XSSFCell) rowOrderNum.createCell(7);

            XSSFCell cell21 = (XSSFCell) rowProject.createCell(3);
            XSSFCell cell22 = (XSSFCell) rowProject.createCell(4);
            XSSFCell cell23 = (XSSFCell) rowProject.createCell(5);
            XSSFCell cell24 = (XSSFCell) rowProject.createCell(6);
            XSSFCell cell25 = (XSSFCell) rowProject.createCell(7);

            XSSFCell cell31 = (XSSFCell) rowCustomer.createCell(3);
            XSSFCell cell32 = (XSSFCell) rowCustomer.createCell(4);
            XSSFCell cell33 = (XSSFCell) rowCustomer.createCell(5);
            XSSFCell cell34 = (XSSFCell) rowCustomer.createCell(6);
            XSSFCell cell35 = (XSSFCell) rowCustomer.createCell(7);

            XSSFCell cell41 = (XSSFCell) rowDateStart.createCell(3);
            XSSFCell cell42 = (XSSFCell) rowDateStart.createCell(4);
            XSSFCell cell43 = (XSSFCell) rowDateStart.createCell(5);
            XSSFCell cell44 = (XSSFCell) rowDateStart.createCell(6);
            XSSFCell cell45 = (XSSFCell) rowDateStart.createCell(7);

            XSSFCell cell51 = (XSSFCell) rowAddress.createCell(3);
            XSSFCell cell52 = (XSSFCell) rowAddress.createCell(4);
            XSSFCell cell53 = (XSSFCell) rowAddress.createCell(5);
            XSSFCell cell54 = (XSSFCell) rowAddress.createCell(6);
            XSSFCell cell55 = (XSSFCell) rowAddress.createCell(7);

            XSSFCell cell61 = (XSSFCell) rowContractor.createCell(3);
            XSSFCell cell62 = (XSSFCell) rowContractor.createCell(4);
            XSSFCell cell63 = (XSSFCell) rowContractor.createCell(5);
            XSSFCell cell64 = (XSSFCell) rowContractor.createCell(6);
            XSSFCell cell65 = (XSSFCell) rowContractor.createCell(7);

            CellStyle cellStyleLeftBold = document.createCellStyle();
            cellStyleLeftBold.setAlignment(HorizontalAlignment.LEFT);
            cellStyleLeftBold.setFont(fontBold);
            cell11.setCellStyle(cellStyleLeftBold);
            cell21.setCellStyle(cellStyleLeftBold);
            cell31.setCellStyle(cellStyleLeftBold);
            cell41.setCellStyle(cellStyleLeftBold);
            cell51.setCellStyle(cellStyleLeftBold);
            cell61.setCellStyle(cellStyleLeftBold);

            CellStyle cellStyleRightBold = document.createCellStyle();
            cellStyleLeftBold.setAlignment(HorizontalAlignment.RIGHT);
            cellStyleLeftBold.setFont(fontBold);
            cell12.setCellStyle(cellStyleRightBold);
            cell22.setCellStyle(cellStyleRightBold);
            cell32.setCellStyle(cellStyleRightBold);
            cell42.setCellStyle(cellStyleRightBold);
            cell52.setCellStyle(cellStyleRightBold);
            cell62.setCellStyle(cellStyleRightBold);

            CellStyle cellStyleCenter = document.createCellStyle();
            cellStyleCenter.setAlignment(HorizontalAlignment.CENTER);

            InputStream inputStream = new FileInputStream(documentFolderPath + "/logo_big.png");
            byte[] imageBytes = IOUtils.toByteArray(inputStream);
            int pictureIdx = document.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
            inputStream.close();
            CreationHelper helper = document.getCreationHelper();
            Drawing drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();

            anchor.setCol1(0);
            anchor.setRow1(1);
            anchor.setRow2(7);
            anchor.setCol2(3);

            drawing.createPicture(anchor, pictureIdx);


            cell11.setCellValue("Заказ №");
            cell12.setCellValue(order.getOrderId() + " от " + order.getDateIncomeStringDDMMYYYY());
            cell21.setCellValue("Проект:");
            cell22.setCellValue("");
            cell31.setCellValue("Клиент:");
            cell32.setCellValue(customer.getFIO());
            cell41.setCellValue("Дата:");
            cell42.setCellValue(order.getDateStartStringDDMMYYYY());
            cell51.setCellValue("Локация:");
            cell52.setCellValue(order.getDeliveryAddress());
            cell61.setCellValue("Подрядчик:");
            cell62.setCellValue("Event Help (098) 646-55-95 / (068) 968-65-65");

            XSSFColor colorBlack = new XSSFColor(Color.BLACK);
            XSSFColor colorLightGray = new XSSFColor(Color.LIGHT_GRAY);

            XSSFCellStyle cellStyleOrderBottomBorderBold = document.createCellStyle();
            cellStyleOrderBottomBorderBold.setAlignment(HorizontalAlignment.LEFT);
            cellStyleOrderBottomBorderBold.setFont(fontBold);
            cellStyleOrderBottomBorderBold.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cellStyleOrderBottomBorderBold.setWrapText(true);
            createBorderB(colorBlack, cellStyleOrderBottomBorderBold);

            XSSFCellStyle cellStyleOrderBottomBorder = document.createCellStyle();
            cellStyleOrderBottomBorder.setAlignment(HorizontalAlignment.LEFT);
            cellStyleOrderBottomBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cellStyleOrderBottomBorder.setWrapText(true);
            createBorderB(colorBlack, cellStyleOrderBottomBorder);


            cell12.setCellStyle(cellStyleOrderBottomBorderBold);
            cell13.setCellStyle(cellStyleOrderBottomBorderBold);
            cell14.setCellStyle(cellStyleOrderBottomBorderBold);
            cell15.setCellStyle(cellStyleOrderBottomBorderBold);

            cell22.setCellStyle(cellStyleOrderBottomBorder);
            cell23.setCellStyle(cellStyleOrderBottomBorder);
            cell24.setCellStyle(cellStyleOrderBottomBorder);
            cell25.setCellStyle(cellStyleOrderBottomBorder);

            cell32.setCellStyle(cellStyleOrderBottomBorder);
            cell33.setCellStyle(cellStyleOrderBottomBorder);
            cell34.setCellStyle(cellStyleOrderBottomBorder);
            cell35.setCellStyle(cellStyleOrderBottomBorder);

            cell42.setCellStyle(cellStyleOrderBottomBorder);
            cell43.setCellStyle(cellStyleOrderBottomBorder);
            cell44.setCellStyle(cellStyleOrderBottomBorder);
            cell45.setCellStyle(cellStyleOrderBottomBorder);

            cell52.setCellStyle(cellStyleOrderBottomBorder);
            cell53.setCellStyle(cellStyleOrderBottomBorder);
            cell54.setCellStyle(cellStyleOrderBottomBorder);
            cell55.setCellStyle(cellStyleOrderBottomBorder);

            cell62.setCellStyle(cellStyleOrderBottomBorder);
            cell63.setCellStyle(cellStyleOrderBottomBorder);
            cell64.setCellStyle(cellStyleOrderBottomBorder);
            cell65.setCellStyle(cellStyleOrderBottomBorder);

            XSSFCellStyle cellHeadStyleCenterBoldGray = document.createCellStyle();
            cellHeadStyleCenterBoldGray.setAlignment(HorizontalAlignment.CENTER);
            cellHeadStyleCenterBoldGray.setFont(fontBold);
            cellHeadStyleCenterBoldGray.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cellHeadStyleCenterBoldGray.setFillForegroundColor(colorLightGray);
            cellHeadStyleCenterBoldGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellHeadStyleCenterBoldGray.setWrapText(true);
            createBorderTRBL(colorBlack, cellHeadStyleCenterBoldGray);

            XSSFCellStyle cellHeadStyleCenterBold = document.createCellStyle();
            cellHeadStyleCenterBold.setAlignment(HorizontalAlignment.CENTER);
            cellHeadStyleCenterBold.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cellHeadStyleCenterBold.setWrapText(true);
            createBorderTRBL(colorBlack, cellHeadStyleCenterBold);

            XSSFCellStyle firstCellStyleLeft = document.createCellStyle();
            firstCellStyleLeft.setAlignment(HorizontalAlignment.LEFT);
            firstCellStyleLeft.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleLeft.setWrapText(true);
            createBorderTRBL(colorBlack, firstCellStyleLeft);

            XSSFCellStyle firstCellStyleRightNoBorder = document.createCellStyle();
            firstCellStyleRightNoBorder.setAlignment(HorizontalAlignment.RIGHT);
            firstCellStyleRightNoBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleRightNoBorder.setWrapText(true);

            XSSFCellStyle firstCellStyleCenterNoBorderBold = document.createCellStyle();
            firstCellStyleCenterNoBorderBold.setAlignment(HorizontalAlignment.CENTER);
            firstCellStyleCenterNoBorderBold.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleCenterNoBorderBold.setFont(fontBold);
            firstCellStyleCenterNoBorderBold.setWrapText(true);

            XSSFCellStyle firstCellStyleCenterNoBorder = document.createCellStyle();
            firstCellStyleCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            firstCellStyleCenterNoBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleCenterNoBorder.setWrapText(true);

            XSSFCellStyle firstCellStyleLeftNoBorder = document.createCellStyle();
            firstCellStyleLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            firstCellStyleLeftNoBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleLeftNoBorder.setWrapText(true);


            CellRangeAddress cellProductName = new CellRangeAddress(9, 9, 2, 3);
            sheet.addMergedRegion(cellProductName);

            List<String> data = new ArrayList<>();
            data.add(0,"№");
            data.add(1,"Категория");
            data.add(2,"Наименование");
            data.add(3,"Кол-во");
            data.add(4,"Цена в день");
            data.add(5,"Период");
            data.add(6,"Сумма");
            data.add(7,"Примечание");

            int rowNum = 9;
            createTableRow2(rowNum, sheet, null, cellHeadStyleCenterBoldGray, data, "head");

            rowNum++;
            int rowCount = 1;
            for(ProductInOrder productInOrder : productsInOrders){
                CellRangeAddress cellProduct = new CellRangeAddress(rowNum, rowNum, 2, 3);
                sheet.addMergedRegion(cellProduct);

                List<String> data1 = new ArrayList<>();
                String model = "";
                if(productInOrder.getProductModel() != null && productInOrder.getProductModel().trim().length() > 0){
                    model = " - (" + productInOrder.getProductModel() + ")";
                }
                data1.add(0,"" + rowCount++);
                data1.add(1,"категория" + productInOrder.getProductId());
                data1.add(2,productInOrder.getProductName() + model);
                data1.add(3,productInOrder.getProductAmountString());
                data1.add(4,productInOrder.getProductPriceString());
                data1.add(5,order.getCountDay() + "");
                data1.add(6,"" + (productInOrder.getProductPrice() * productInOrder.getProductAmount() * order.getCountDay()));
                data1.add(7,order.getDescription());
                createTableRow2(rowNum, sheet, firstCellStyleLeft, cellHeadStyleCenterBold, data1, null);

                rowNum++;
            }

            Row rowSum = sheet.createRow(++rowNum);

            rowNum++;
            Row rowDiscount = sheet.createRow(++rowNum);

            rowNum++;
            Row rowDelivery = sheet.createRow(++rowNum);

            rowNum++;
            Row rowItog = sheet.createRow(++rowNum);

            rowNum+=2;
            Row rowFOP = sheet.createRow(++rowNum);

            rowNum++;
            Row rowFaximile = sheet.createRow(++rowNum);

            XSSFCell cellSumLable = (XSSFCell) rowSum.createCell(5);
            XSSFCell cellSum = (XSSFCell) rowSum.createCell(7);
            cellSumLable.setCellStyle(firstCellStyleCenterNoBorderBold);
            cellSum.setCellStyle(cellHeadStyleCenterBold);
            cellSumLable.setCellValue("Итого");
            cellSum.setCellValue(order.getPriceDiscountString() + " грн.");

            XSSFCell cellDiscountLable = (XSSFCell) rowDiscount.createCell(5);
            XSSFCell cellDiscount = (XSSFCell) rowDiscount.createCell(7);
            cellDiscountLable.setCellStyle(firstCellStyleCenterNoBorderBold);
            cellDiscount.setCellStyle(cellHeadStyleCenterBold);
            cellDiscountLable.setCellValue("Скидка");
            cellDiscount.setCellValue(0.0 + " грн.");

            XSSFCell cellDeliveryLable = (XSSFCell) rowDelivery.createCell(5);
            XSSFCell cellDelivery = (XSSFCell) rowDelivery.createCell(7);
            cellDeliveryLable.setCellStyle(firstCellStyleCenterNoBorderBold);
            cellDelivery.setCellStyle(cellHeadStyleCenterBold);
            cellDeliveryLable.setCellValue("Доставка");
            cellDelivery.setCellValue(0.0 + " грн.");

            XSSFCell cellItogLable = (XSSFCell) rowItog.createCell(5);
            XSSFCell cellItog = (XSSFCell) rowItog.createCell(7);
            cellItogLable.setCellStyle(firstCellStyleCenterNoBorderBold);
            cellItog.setCellStyle(cellHeadStyleCenterBold);
            cellItogLable.setCellValue("Всего");
            cellItog.setCellValue(order.getPriceDiscountString() + " грн.");

            XSSFCell cellFOPLable = (XSSFCell) rowFOP.createCell(5);
            XSSFCell cellFOP = (XSSFCell) rowFOP.createCell(7);
            cellFOPLable.setCellStyle(firstCellStyleCenterNoBorderBold);
            cellFOP.setCellStyle(cellHeadStyleCenterBold);
            cellFOPLable.setCellValue("Безнал ФОП");
            cellFOP.setCellValue(0.0 + " грн.");

            XSSFCell cellFaksimileLable = (XSSFCell) rowFaximile.createCell(5);
            XSSFCell cellFaksimile = (XSSFCell) rowFaximile.createCell(7);
            cellFaksimileLable.setCellStyle(firstCellStyleCenterNoBorder);
            cellFaksimile.setCellStyle(firstCellStyleCenterNoBorder);
            cellFaksimileLable.setCellValue("______________");
            cellFaksimile.setCellValue("/ Креденцер В.А./");

            document.write(fileOutputStream);
        }catch (IOException ex){

        }finally {
            try {
                fileOutputStream.flush();
            } catch (IOException e) {
                e.printStackTrace();
            }finally {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return resultFile;
    }

    /*
    По принцыпу сметы. Список оборудки сумма, адрес доставки, имя заказчика.
     */
    public static File createXlsX(String filePath, Order order, List<ProductInOrder> productsInOrders, Customer customer){
        File resultFile = new File(filePath);
        XSSFWorkbook document = new XSSFWorkbook();
        FileOutputStream fileOutputStream = null;

        try {
            fileOutputStream = new FileOutputStream(resultFile);
            XSSFSheet sheet = document.createSheet("order_info");
            Row rowOrderNum = sheet.createRow(1);
            Row rowCustomer = sheet.createRow(2);
            Row rowAddress = sheet.createRow(3);
            Row rowShipOut = sheet.createRow(4);
            Row rowShipBack = sheet.createRow(5);
            CellRangeAddress cellRangeOrderNum = new CellRangeAddress(1, 1, 1, 2);
            CellRangeAddress cellRangeFIO = new CellRangeAddress(2, 2, 1, 5);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(3, 3, 1, 5);
            sheet.addMergedRegion(cellRangeOrderNum);
            sheet.addMergedRegion(cellRangeFIO);
            sheet.addMergedRegion(cellRangeAddress);
            sheet.setColumnWidth(0, 6000);
            sheet.setColumnWidth(1, 6000);
            sheet.setColumnWidth(2, 3000);
            sheet.setColumnWidth(3, 3000);
            sheet.setColumnWidth(4, 3000);
            sheet.setColumnWidth(5, 3000);
            Font fontBold = document.createFont();
            fontBold.setBold(true);

            XSSFCell cell11 = (XSSFCell) rowOrderNum.createCell(1);
            XSSFCell cell12 = (XSSFCell) rowOrderNum.createCell(2);
            XSSFCell cell13 = (XSSFCell) rowOrderNum.createCell(3);
            XSSFCell cell14 = (XSSFCell) rowOrderNum.createCell(4);

            CellStyle cellStyleLeftBold = document.createCellStyle();
            cellStyleLeftBold.setAlignment(HorizontalAlignment.CENTER);
            cellStyleLeftBold.setFont(fontBold);
            cell11.setCellStyle(cellStyleLeftBold);
            cell12.setCellStyle(cellStyleLeftBold);

            CellStyle cellStyleCenter = document.createCellStyle();
            cellStyleCenter.setAlignment(HorizontalAlignment.CENTER);
            cell13.setCellStyle(cellStyleCenter);
            cell14.setCellStyle(cellStyleCenter);

            cell11.setCellValue("Заказ №" + order.getOrderId() + " от " + order.getDateIncomeStringDDMMYYYY());

            XSSFCell cell21 = ((XSSFRow) rowCustomer).createCell(0);
            XSSFCell cell22 = ((XSSFRow) rowCustomer).createCell(1);
            cell21.setCellValue("Заказчик:");
            cell22.setCellValue(customer.getFIO());

            XSSFCell cell31 = ((XSSFRow) rowAddress).createCell(0);
            XSSFCell cell32 = ((XSSFRow) rowAddress).createCell(1);
            cell31.setCellValue("Адрес доставки:");
            cell32.setCellValue(order.getDeliveryAddress());

            XSSFCell cell41 = ((XSSFRow) rowShipOut).createCell(0);
            XSSFCell cell42 = ((XSSFRow) rowShipOut).createCell(1);
            cell41.setCellValue("Дата отгрузки:");
            cell42.setCellValue(order.getDateShipOutStringDDMMYYYY());

            XSSFCell cell51 = ((XSSFRow) rowShipBack).createCell(0);
            XSSFCell cell52 = ((XSSFRow) rowShipBack).createCell(1);
            cell51.setCellValue("Дата возврата:");
            cell52.setCellValue(order.getDateShipBackStringDDMMYYYY());

            XSSFColor colorBlack = new XSSFColor(Color.BLACK);
            XSSFColor colorLightGray = new XSSFColor(Color.LIGHT_GRAY);

            XSSFCellStyle cellHeadStyleCenterBoldGray = document.createCellStyle();
            cellHeadStyleCenterBoldGray.setAlignment(HorizontalAlignment.CENTER);
            cellHeadStyleCenterBoldGray.setFont(fontBold);
            cellHeadStyleCenterBoldGray.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cellHeadStyleCenterBoldGray.setFillForegroundColor(colorLightGray);
            cellHeadStyleCenterBoldGray.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cellHeadStyleCenterBoldGray.setWrapText(true);
            createBorderTRBL(colorBlack, cellHeadStyleCenterBoldGray);

            XSSFCellStyle cellHeadStyleCenterBold = document.createCellStyle();
            cellHeadStyleCenterBold.setAlignment(HorizontalAlignment.CENTER);
            cellHeadStyleCenterBold.setVerticalAlignment(VerticalAlignment.BOTTOM);
            cellHeadStyleCenterBold.setWrapText(true);
            createBorderTRBL(colorBlack, cellHeadStyleCenterBold);

            XSSFCellStyle firstCellStyleLeft = document.createCellStyle();
            firstCellStyleLeft.setAlignment(HorizontalAlignment.LEFT);
            firstCellStyleLeft.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleLeft.setWrapText(true);
            createBorderTRBL(colorBlack, firstCellStyleLeft);

            XSSFCellStyle firstCellStyleRightNoBorder = document.createCellStyle();
            firstCellStyleRightNoBorder.setAlignment(HorizontalAlignment.RIGHT);
            firstCellStyleRightNoBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleRightNoBorder.setWrapText(true);

            XSSFCellStyle firstCellStyleCenterNoBorder = document.createCellStyle();
            firstCellStyleCenterNoBorder.setAlignment(HorizontalAlignment.CENTER);
            firstCellStyleCenterNoBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleCenterNoBorder.setFont(fontBold);
            firstCellStyleCenterNoBorder.setWrapText(true);

            XSSFCellStyle firstCellStyleLeftNoBorder = document.createCellStyle();
            firstCellStyleLeftNoBorder.setAlignment(HorizontalAlignment.LEFT);
            firstCellStyleLeftNoBorder.setVerticalAlignment(VerticalAlignment.BOTTOM);
            firstCellStyleLeftNoBorder.setWrapText(true);

            CellRangeAddress cellFirstColl = new CellRangeAddress(6, 6, 0, 1);
            sheet.addMergedRegion(cellFirstColl);

            List<String> data = new ArrayList<>();
            data.add(0,"Наименование");
            data.add(1,"Количество");
            data.add(2,"Цена в день");
            data.add(3,"Количество дней");
            data.add(4,"Сумма");
            createTableRow(6, sheet, null, cellHeadStyleCenterBoldGray, data, "head");

            int rowNum = 7;
            for(ProductInOrder productInOrder : productsInOrders){
                CellRangeAddress cellFirst = new CellRangeAddress(rowNum, rowNum, 0, 1);
                sheet.addMergedRegion(cellFirst);

                List<String> data1 = new ArrayList<>();
                String model = "";
                if(productInOrder.getProductModel() != null && productInOrder.getProductModel().trim().length() > 0){
                    model = " - (" + productInOrder.getProductModel() + ")";
                }
                data1.add(0,productInOrder.getProductName() + model);
                data1.add(1,productInOrder.getProductAmountString());
                data1.add(2,productInOrder.getProductPriceString());
                data1.add(3,order.getCountDay() + "");
                data1.add(4,"" + (productInOrder.getProductPrice() * productInOrder.getProductAmount() * order.getCountDay()));
                createTableRow(rowNum, sheet, firstCellStyleLeft, cellHeadStyleCenterBold, data1, null);

                rowNum++;
            }

            Row rowSum = sheet.createRow(rowNum);


            XSSFCell cellItog = (XSSFCell) rowSum.createCell(4);
            cellItog.setCellStyle(firstCellStyleRightNoBorder);
            XSSFCell cellSum = (XSSFCell) rowSum.createCell(5);
            cellSum.setCellStyle(firstCellStyleCenterNoBorder);
            XSSFCell cellGrn = (XSSFCell) rowSum.createCell(6);
            cellGrn.setCellStyle(firstCellStyleLeftNoBorder);
            cellItog.setCellValue("ИТОГО:");
            cellSum.setCellValue(order.getPriceDiscountString());
            cellGrn.setCellValue("грн.");

            document.write(fileOutputStream);
        }catch (IOException ex){

        }finally {
            try {
                fileOutputStream.flush();
            } catch (IOException e) {
                e.printStackTrace();
            }finally {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        return resultFile;
    }

    private static void createTableRow2(int rowNum, XSSFSheet sheet, XSSFCellStyle firstCellStyle, XSSFCellStyle cellStyle, List<String> data, String isHead){
        Row row = sheet.createRow(rowNum);


        XSSFCell cell1 = (XSSFCell) row.createCell(0);
        cell1.setCellStyle(cellStyle);
        XSSFCell cell2 = (XSSFCell) row.createCell(1);
        cell2.setCellStyle(cellStyle);
        XSSFCell cell3 = (XSSFCell) row.createCell(2);
        cell3.setCellStyle(isHead != null ? cellStyle : firstCellStyle);
        XSSFCell _cell3 = (XSSFCell) row.createCell(3);
        _cell3.setCellStyle(isHead != null ? cellStyle : firstCellStyle);
        XSSFCell cell4 = (XSSFCell) row.createCell(4);
        cell4.setCellStyle(cellStyle);
        XSSFCell cell5 = (XSSFCell) row.createCell(5);
        cell5.setCellStyle(cellStyle);
        XSSFCell cell6 = (XSSFCell) row.createCell(6);
        cell6.setCellStyle(cellStyle);
        XSSFCell cell7 = (XSSFCell) row.createCell(7);
        cell7.setCellStyle(cellStyle);
        XSSFCell cell8 = (XSSFCell) row.createCell(8);
        cell8.setCellStyle(cellStyle);

        if(isHead != null) {
            cell1.setCellValue(data.get(0));
            cell2.setCellValue(data.get(1));
            cell3.setCellValue(data.get(2));
            cell4.setCellValue(data.get(3));
            cell5.setCellValue(data.get(4));
            cell6.setCellValue(data.get(5));
            cell7.setCellValue(data.get(6));
            cell8.setCellValue(data.get(7));
        }else {
            cell1.setCellValue(Integer.parseInt(data.get(0)));
            cell2.setCellValue(data.get(1));
            cell3.setCellValue(data.get(2));
            cell4.setCellValue(Integer.parseInt(data.get(3)));
            cell5.setCellValue(Double.parseDouble(data.get(4)) + " грн.");
            cell6.setCellValue(Integer.parseInt(data.get(5)));
            cell7.setCellValue(Double.parseDouble(data.get(6)) + " грн.");
            cell8.setCellValue(data.get(7));
        }
    }

    private static void createTableRow(int rowNum, XSSFSheet sheet, XSSFCellStyle firstCellStyle, XSSFCellStyle cellStyle, List<String> data, String isHead){
        Row row = sheet.createRow(rowNum);


        XSSFCell cell1 = (XSSFCell) row.createCell(0);
        cell1.setCellStyle(isHead != null ? cellStyle : firstCellStyle);
        XSSFCell _cell1 = (XSSFCell) row.createCell(1);
        _cell1.setCellStyle(isHead != null ? cellStyle : firstCellStyle);
        XSSFCell cell2 = (XSSFCell) row.createCell(2);
        cell2.setCellStyle(cellStyle);
        XSSFCell cell3 = (XSSFCell) row.createCell(3);
        cell3.setCellStyle(cellStyle);
        XSSFCell cell4 = (XSSFCell) row.createCell(4);
        cell4.setCellStyle(cellStyle);
        XSSFCell cell5 = (XSSFCell) row.createCell(5);
        cell5.setCellStyle(cellStyle);

        cell1.setCellValue(data.get(0));
        if(isHead != null) {
            cell2.setCellValue(data.get(1));
            cell3.setCellValue(data.get(2));
            cell4.setCellValue(data.get(3));
            cell5.setCellValue(data.get(4));
        }else {
            cell2.setCellValue(Integer.parseInt(data.get(1)));
            cell3.setCellValue(data.get(2));
            cell4.setCellValue(Integer.parseInt(data.get(3)));
            cell5.setCellValue(data.get(4));
        }
    }

    private static void createBorderTRBL(XSSFColor color, XSSFCellStyle cellStyle) {
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.TOP, color);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.LEFT, color);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.RIGHT, color);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, color);
    }
    private static void createBorderB(XSSFColor color, XSSFCellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderColor(XSSFCellBorder.BorderSide.BOTTOM, color);
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
            String model = "";
            if(productsInOrder.get(i-1).getProductModel() != null && productsInOrder.get(i-1).getProductModel().trim().length() > 0){
                model = " - (" + productsInOrder.get(i-1).getProductModel() + ")";
            }
            table.getRow(i).getCell(0).setText(productsInOrder.get(i-1).getProductName() + model);
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
