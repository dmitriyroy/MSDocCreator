import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.hwpf.usermodel.CharacterRun;

import javax.xml.bind.JAXBElement;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Main_0 {

    public static void main(String[] args) throws IOException {


        String filePathDocTemplateX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.docx";
        String filePathDocX ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.docx";
        String filePathDocTemplate ="D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/dogovor_fiz_template.doc";
        String filePathDoc = "D:/Java/t2studio/CRM_eventhelp_web/src/main/resources/static/documents/Fiz_Document.doc";

/*
        POIFSFileSystem fs = null;
        try {
            fs = new POIFSFileSystem(new FileInputStream(filePathDocTemplate));
            HWPFDocument doc = new HWPFDocument(fs);
            doc = replaceText(doc, "ORDER_NUM", 777);
            doc = replaceText(doc, "FIO", "Рой Дмитрий Николаевич");
            doc = replaceText(doc, "ORDER_SUM", 156.35);
            doc = replaceText(doc, "DAY1", 25);
            doc = replaceText(doc, "DAY2", 27);
            doc = replaceText(doc, "DAY3", 29);
            doc = replaceText(doc, "MONTH1", 12);
            doc = replaceText(doc, "MONTH2", 12);
            doc = replaceText(doc, "MONTH3", 12);
            doc = replaceText(doc, "YEAR1", 2018);
            doc = replaceText(doc, "YEAR2", 2018);
            doc = replaceText(doc, "YEAR3", 2018);
            doc = replaceText(doc, "SNAME", "Рой");
            doc = replaceText(doc, "FNAME", "Дмитрий");
            doc = replaceText(doc, "PNAME", "Николаевич");
            saveWord(filePathDoc, doc);
        }
        catch(FileNotFoundException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }finally{
            if(fs != null) {
                fs.close();
            }
        }

*/

        FileInputStream fis = new FileInputStream(filePathDocTemplateX);
        XWPFDocument fileShablon = new XWPFDocument(fis);
        List<XWPFParagraph> paragraphs = fileShablon.getParagraphs();
        XWPFWordExtractor extractor = new XWPFWordExtractor(fileShablon);
        String paragraphText = extractor.getText();
//        System.out.println("paragraphText = " + paragraphText);

        File outFile = new File(filePathDocX);
        XWPFDocument document = new XWPFDocument();
        FileOutputStream fos = new FileOutputStream(outFile);

        int i = 0;
        for(XWPFParagraph paragraph1 : paragraphs){
            i++;
            XWPFParagraph paragraph = document.createParagraph();
            System.out.println(i + " - paragraph1.getText() : " + paragraph1.getText());
            System.out.println(i + " - paragraph1.getAlignment() : " + paragraph1.getAlignment());
            System.out.println(i + " - paragraph1.getFontAlignment() : " + paragraph1.getFontAlignment());
            System.out.println(i + " - paragraph1.getNumID() : " + paragraph1.getNumID());
            System.out.println(i + " - paragraph1.getNumIlvl() : " + paragraph1.getNumIlvl());
            System.out.println(i + " - paragraph1.getNumLevelText() : " + paragraph1.getNumLevelText());
            System.out.println(i + " - paragraph1.getStyle() : " + paragraph1.getStyle());
            paragraph.setAlignment(paragraph1.getAlignment());
//            System.out.println(i + " - paragraph.getText() : " + paragraph.getText());
//            document.setParagraph(paragraph1, i++);

//            paragraph.getText();
//            paragraph.getParagraphText();
//            paragraph.getAlignment();
//            System.out.println("paragraph.getSpacingBefore() = " + paragraph.getSpacingBefore());
//            System.out.println("paragraph.getText() : " + paragraph.getText());
//            System.out.println("paragraph.getParagraphText() : " + paragraph.getParagraphText());
//            System.out.println("paragraph.getSpacingAfter() = " + paragraph.getSpacingAfter());
        }
//        document.write(fos);
//        fos.flush();
//        fos.close();

//        XWPFParagraph paragraph = document.createParagraph();
//        paragraph.setAlignment(ParagraphAlignment.CENTER);
//        XWPFRun run = paragraph.createRun();
//        for(String row : rows) {
//            run.setText(row);
//            run.addTab();
//            run.addBreak();
//        }



//        createDocx(filePathDocx, fileRows);

//        List<String> fileRows = new ArrayList<>();
//        fileRows.add("Название\n");
//        fileRows.add("Строка 1\n");
//        fileRows.add("Строка 2\n");
//        fileRows.add("Строка 3\n");
//        fileRows.add("Строка 4\n");

//        createDocx(filePathDocx, fileRows);
//        createDocx(filePathDoc, fileRows);

//        String filePathXlsx = "D:/xwpfDocument.xlsx";
//        createXlsx(filePathXlsx, fileRows);

    }

//    private WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
//        WordprocessingMLPackage template = WordprocessingMLPackage.load(new FileInputStream(new File(name)));
//        return template;
//    }
//    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
//        List<Object> result = new ArrayList<Object>();
//        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();
//
//        if (obj.getClass().equals(toSearch))
//            result.add(obj);
//        else if (obj instanceof ContentAccessor) {
//            List<?> children = ((ContentAccessor) obj).getContent();
//            for (Object child : children) {
//                result.addAll(getAllElementFromObject(child, toSearch));
//            }
//
//        }
//        return result;
//    }
//
//    private void replacePlaceholder(WordprocessingMLPackage template, String name, String placeholder ) {
//        List<Object> texts = getAllElementFromObject(template.getMainDocumentPart(), Text.class);
//
//        for (Object text : texts) {
//            Text textElement = (Text) text;
//            if (textElement.getValue().equals(placeholder)) {
//                textElement.setValue(name);
//            }
//        }
//    }
//
//    private void writeDocxToStream(WordprocessingMLPackage template, String target) throws IOException, Docx4JException {
//        File f = new File(target);
//        template.save(f);
//    }


    private static HWPFDocument replaceText(HWPFDocument doc, String findText, Object replaceText){
        Range r1 = doc.getRange();

        for (int i = 0; i < r1.numSections(); ++i ) {
            Section s = r1.getSection(i);
            for (int x = 0; x < s.numParagraphs(); x++) {
                Paragraph p = s.getParagraph(x);
                for (int z = 0; z < p.numCharacterRuns(); z++) {
                    CharacterRun run = p.getCharacterRun(z);
                    String text = run.text();
//                    if(text.contains(findText.toLowerCase())) {
                        run.replaceText(findText, replaceText.toString());
//                    }
//                    run.replaceText("ORDER_NUM", 777);
//                    run.replaceText("FIO", "Рой Дмитрий Николаевич");
//                    run.replaceText("ORDER_SUM", 156.35);
//                    run.replaceText("DAY1", 25);
//                    run.replaceText("DAY2", 27);
//                    run.replaceText("DAY3", 29);
//                    run.replaceText("MONTH1", 12);
//                    run.replaceText("MONTH2", 12);
//                    run.replaceText("MONTH3", 12);
//                    run.replaceText("YEAR1", 2018);
//                    run.replaceText("YEAR2", 2018);
//                    run.replaceText("YEAR3", 2018);
//                    run.replaceText("SNAME", "Рой");
//                    run.replaceText("FNAME", "Дмитрий");
//                    run.replaceText("PNAME", "Николаевич");
                }
            }
        }
        return doc;
    }

    private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(filePath);
            doc.write(out);
        }
        finally{
            out.close();
        }
    }
    public static File createDocx(String filePath, List<String> rows) throws IOException {
        File outFile = new File(filePath);
        XWPFDocument document = new XWPFDocument();
        FileOutputStream fileOutputStream = new FileOutputStream(outFile);

        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        for(String row : rows) {
            run.setText(row);
            run.addTab();
            run.addBreak();
        }
        document.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();

        return outFile;
    }

    public static File createXlsx(String filePath, List<String> rows) throws IOException {
        File outFile = new File(filePath);
        XSSFWorkbook document = new XSSFWorkbook();
        FileOutputStream fileOutputStream = new FileOutputStream(outFile);

        XSSFSheet sheet = document.createSheet("DimaTestSheet");
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("Вася");
        row.createCell(1).setCellValue(new SimpleDateFormat("yyyy-MM-dd hh:mm:ss").format(new Date()));
        row.createCell(2).setCellValue(new Date());

        document.write(fileOutputStream);
        fileOutputStream.flush();
        fileOutputStream.close();

        return outFile;
    }
}
