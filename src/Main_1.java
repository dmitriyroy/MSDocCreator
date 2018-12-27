import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Main_1 {

    public static void main(String[] args) {
        String filePathDocTemplateX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.docx";
        String filePathDocX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.docx";
//        String filePathDocTemplate ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.doc";
//        String filePathDoc = "D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.doc";

        try {
            createDocX(filePathDocTemplateX, filePathDocX, createDataMap());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static File createDocX(String templateFilePath, String resultFilePath, Map<String, Object> replaceData) throws IOException {
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
                paragraph.setFontAlignment(paragraph1.getFontAlignment());
                paragraph.setNumID(paragraph1.getNumID());
                paragraph.setSpacingAfter(paragraph1.getSpacingAfter());
                paragraph.setSpacingBefore(paragraph1.getSpacingBefore());
                paragraph.setVerticalAlignment(paragraph1.getVerticalAlignment());
                run.setText(replaceDataValue(text, replaceData));
//            paragraph.setStyle(paragraph1.getStyle());
//            paragraph.setSpacingAfterLines(paragraph1.getSpacingAfterLines());
//            paragraph.setSpacingBeforeLines(paragraph1.getSpacingBeforeLines());
//            paragraph.setFirstLineIndent(paragraph1.getFirstLineIndent());
//            paragraph.setSpacingLineRule(paragraph1.getSpacingLineRule());
//            paragraph.setSpacingBetween(paragraph1.getSpacingBetween());
            }else{
                createTable(document);
            }
        }

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

    private static void createTable(XWPFDocument document){
//create table
        XWPFTable table = document.createTable();
        //create first row
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("col one, row one");
        tableRowOne.addNewTableCell().setText("col two, row one");
        tableRowOne.addNewTableCell().setText("col three, row one");
        //create second row
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("col one, row two");
        tableRowTwo.getCell(1).setText("col two, row two");
        tableRowTwo.getCell(2).setText("col three, row two");
        //create third row
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("col one, row three");
        tableRowThree.getCell(1).setText("col two, row three");
        tableRowThree.getCell(2).setText("col three, row three");
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
}
