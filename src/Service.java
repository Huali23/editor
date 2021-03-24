import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

public class Service {

    private XSSFCellStyle styleBorder;
    private XSSFCellStyle styleYellow;
    private XSSFFont fontEM;
    private XSSFFont fontNormal;
    private final String[] EULongArr = new String[]{"Austria", "Belgium", "Bulgaria", "Croatia", "Cyprus", "Czech Republic", "Denmark", "England", "Estonia", "Finland", "France", "Germany", "Greece", "Hungary", "Iceland", "Ireland", "Italy", "Latvia ", "Liechtenstein", "Lithuania", "Luxembourg", "Malta", "Netherlands", "North Ireland", "Norway", "Poland", "Portugal", "Romania", "Scotland", "Slovakia", "Slovenia", "Spain", "Sweden", "Switzerland"};
    private final String[] EUShortArr = new String[]{"at","be","bg","hr","cy","cs","dk","uk","ee","fi","fr","de","gr","hu","is","ie","it","lv","li","lt","lu","mt","nl","ie","no","pl","pt","ro","sk","si","es","se","ch"};

    private static int i;

    private static int j;

    private final String[] editorKey = new String[]{"@Name", "@EMail", "@Reference", "@Link", "@Google Scholar", "@TC", "@PY", "@Affiliation"};

    public Service() {

    }


    public void saveFile(XSSFWorkbook workbook) {
        //弹出文件选择框
        JFileChooser chooser = new JFileChooser();

        Date date = new Date();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
        String dateStr = dateFormat.format(date);
        chooser.setSelectedFile(new File("editor-"+ dateStr + ".xlsx"));
        int option = chooser.showSaveDialog(null);
        if(option==JFileChooser.APPROVE_OPTION){
            File file = chooser.getSelectedFile();
            try {
                FileOutputStream fos = new FileOutputStream(file);
                workbook.write(fos);
                fos.close();
            } catch (IOException e) {
                System.err.println("IO异常");
                e.printStackTrace();
            }
        }
    }



    public XSSFWorkbook changeToExcel(File[] files) {
        //生成EXCEL模板
        XSSFWorkbook workbook = this.createExcelFile();
        InputStream is = null;
        Scanner scanner = null;
        // 初始化sheet1 sheet2 行索引
        i = 1;
        j = 1;
        try {
            for (File file : files) {
                is = new FileInputStream(file);
                scanner = new Scanner(is);
                if (scanner.hasNext()) {
                    scanner.nextLine();
                }
                while (scanner.hasNext()) {
                    String text = scanner.nextLine();
                    String[] words = text.split("\t");
                    this.valueFactory(workbook, words);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }
    private void createRows(XSSFSheet sheet, List<String> valueList, Integer index, String style) {
        XSSFRow row = sheet.createRow(index);
        for (int k = 0; k < valueList.size(); k++) {
            this.createCells(row, k, valueList.get(k), style);
        }
    }

    private void createRows(XSSFSheet sheet, List<String> valueList, Integer index) {
        this.createRows(sheet, valueList, index, null);
    }
    private void valueFactory(XSSFWorkbook workbook, String[] words) {
        if (StringUtils.isBlank(words[24])) {
            return;
        }
        String[] names = null;
        String[] emails = null;
        List<String> valueList = new ArrayList<>();
        valueList.add(words[1]);  // 0 name
        valueList.add(words[24]); // 1 email
        StringBuilder reference = new StringBuilder();
        reference.append(words[8]).append(".  ").append(words[42]).append("  ").append(words[44]);
        if (StringUtils.isNotBlank(words[45])){
            reference.append(",  ").append(words[45]);
        }
        if (StringUtils.isNotBlank(words[51])) { //页码
            reference.append(",  ").append(words[51]);
        }
        if (StringUtils.isNotBlank(words[52])){
            reference.append("-").append(words[52]);
        }
        reference.append(".");
        valueList.add(reference.toString()); // 2 reference
        valueList.add("http://dx.doi.org/" + words[54]); // 3 link
        valueList.add("http://scholar.google.com/scholar?q=" + words[24]); // 4 googleScholar
        valueList.add(words[31]); // 5 tc
        valueList.add(words[44]); // 6 py
        valueList.add(words[23]); // 7 affiliation
        if (valueList.get(0).contains(";")) {  // 拆分多个姓名， 若1人，names为空
            names = valueList.get(0).split("; ");
        }
        if (valueList.get(1).contains(";")) {
            emails = valueList.get(1).split("; ");
        }
        XSSFSheet sheet1 = workbook.getSheet("sheet1");
        XSSFSheet sheet2 = workbook.getSheet("sheet2");
        List<String[]> matchList = null;
        if (names == null) { //一对一
            String color = getCellColor(valueList.get(7), valueList.get(1));
            this.createRows(sheet1, valueList, i, color);
            i++;
        } else {
            if (emails == null) { //多对一
                matchList = this.matchNameAndEmail(names, new String[]{valueList.get(1)});
            } else {              //多对多
                if (names.length == emails.length) {   // 相等
                    for (int k = 0; k < names.length; k++) {
                        valueList.set(0, names[k]);
                        valueList.set(1, emails[k]);
                        valueList.set(4, "http://scholar.google.com/scholar?q=" + emails[k]);
                        String color = getCellColor(valueList.get(7), emails[k]);
                        this.createRows(sheet1, valueList, i, color);
                        i++;
                    }
                    return;
                } else { // 不等
                    matchList = this.matchNameAndEmail(names, emails);
                }
            }
            // 智能识别姓名和邮箱  匹配的放在sheet1
            if (matchList != null && matchList.size() > 0) {
                for (String[] match : matchList) {
                    valueList.set(0, match[0]);
                    valueList.set(1, match[1]);
                    valueList.set(4, "http://scholar.google.com/scholar?q=" + match[1]);
                    String color = getCellColor(valueList.get(7), match[1]);
                    this.createRows(sheet1, valueList, i, color);
                    i++;
                }
            } else { // 不匹配 放在sheet2
                String color = getCellColor(valueList.get(7), valueList.get(1));
                this.createRows(sheet2, valueList, j, color);
                j++;
            }
        }
    }

    private String getCellColor(String affiliation, String mails) {
        if (this.isEuropean(affiliation, mails)) {
            return "yellow";
        } else {
            return "";
        }

    }

    private boolean isEuropean(String affiliation, String mails) {
        if (mails.contains(".")) {
            String subMails = mails.substring(mails.lastIndexOf("."));
            for (String str: EULongArr) {
                if (affiliation.contains(str)) {
                    return true;
                }
            }
            for (String str: EUShortArr) {
                if (str.equals(subMails)) {
                    return true;
                }
            }
            return false;
        } else {
            return true;
        }
    }


    private List<String[]> matchNameAndEmail(String[] names, String[] emails) {
        List<String[]> matchList = new ArrayList<>();
        for (String name : names) { //循环所有姓名
            String lowName = name.toLowerCase(); //转小写
            if (lowName.contains(",")) {
                String[] splitName = lowName.split(", "); // name 转小写之后拆分firstName secondName
                for (int k = 0; k < emails.length; k++) {  // 双层循环逐次比较mail
                    if (StringUtils.isNotBlank(emails[k])) {
                        String lowEmail = emails[k].toLowerCase();
                        // 邮箱包含 姓名组合
                        if ((splitName.length > 1)&&((lowEmail.contains(splitName[0] + splitName[1])) || (lowEmail.contains(splitName[1] + splitName[0])))) {
                            matchList.add(new String[]{name, emails[k]});
                            emails[k] = ""; // 置空，下次不循环
                            break;
                        }
                        // firstName 长度大于5  肯定是外国人，否则可能是中国人
                        String subEmail = emails[k].split("@")[0];
                        boolean result;
                        if (splitName[0].length() > 4) {
                            result = this.isSubSequence(subEmail, splitName[0]);
                        } else {
                            result = this.isSubSequence(subEmail, splitName[0] + splitName[1]) || this.isSubSequence(subEmail, splitName[1] + splitName[0]);
                        }
                        if (result) {
                            matchList.add(new String[]{name, emails[k]});
                            emails[k] = ""; // 置空，下次不循环
                            break;
                        }
                    }
                }
            }

        }
        return matchList;
    }

    private boolean isSubSequence(String subEmail, String subName) {
        int len1 = 0;
        int len2 = 0;
        if (subEmail.length() >= subName.length()) {
            while (len1 < subEmail.length() && len2 < subName.length()) {
                if (subEmail.charAt(len1) == subName.charAt(len2)) {
                    len1++;
                    len2++;
                } else {
                    len1++;
                }
            }
            return len2 == subName.length();
        } else {
            return false;
        }
    }

    // 创建单元格
    private void createCells(XSSFRow row, int cellIndex, String text, String style) {
        XSSFCell cell = row.createCell(cellIndex);
        if (StringUtils.isNotBlank(text)) {
            if (2 == cellIndex) {
                XSSFRichTextString richText = new XSSFRichTextString(text);
                richText.applyFont(this.fontNormal);
                richText.applyFont(text.indexOf("."), text.substring(0,text.length() - 1).lastIndexOf("."), fontEM);
                cell.setCellValue(richText);
            } else {
                cell.setCellValue(text);
            }
        } else {
            cell.setCellValue("");
        }
        if ("".equals(style)) {
            cell.setCellStyle(styleBorder);
        } else {
            cell.setCellStyle(styleYellow);
        }
    }

    private XSSFWorkbook createExcelFile() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        //边框
        styleBorder = workbook.createCellStyle();
        styleBorder.setBorderTop(HSSFCellStyle.BORDER_THIN);
        styleBorder.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        styleBorder.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        styleBorder.setBorderRight(HSSFCellStyle.BORDER_THIN);
        //正常字体
        this.fontNormal= workbook.createFont();
        this.fontNormal.setFontName("Calibri");
        styleBorder.setFont(this.fontNormal);
        //黄色
        styleYellow = workbook.createCellStyle();
        styleYellow.cloneStyleFrom(styleBorder);
        styleYellow.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        styleYellow.setFillBackgroundColor(IndexedColors.GOLD.getIndex());
        styleYellow.setFillForegroundColor(IndexedColors.GOLD.getIndex());
        //斜体字体
        this.fontEM = workbook.createFont();
        this.fontEM.setItalic(true);

        //生成第一行标题行
        workbook = this.createFirstRow(workbook, "sheet1");
        workbook = this.createFirstRow(workbook, "sheet2");
        return workbook;
    }

    private XSSFWorkbook createFirstRow(XSSFWorkbook workbook, String sheetName) {
        XSSFSheet sheet = workbook.createSheet(sheetName);
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell;
        int i = 0;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("Name");
        sheet.setColumnWidth(i, 4000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("EMail");
        sheet.setColumnWidth(i, 8000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("Reference");
        sheet.setColumnWidth(i, 15000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("Link");
        sheet.setColumnWidth(i, 8000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("Google Scholar");
        sheet.setColumnWidth(i, 8000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("TC");
        sheet.setColumnWidth(i, 2000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("PY");
        sheet.setColumnWidth(i, 2000);
        i++;
        cell = row.createCell(i);
        cell.setCellStyle(styleBorder);
        cell.setCellValue("Affiliation");
        sheet.setColumnWidth(i, 12000);
        return workbook;
    }

    public void directoryToList(File file, List<File> fileList) {
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            if (files != null) {
                for (File file1 : files) {
                    directoryToList(file1, fileList);
                }
            }
        } else {
            String fileName = file.getName();
            String suffix = fileName.substring(fileName.lastIndexOf(".") + 1);
            if ("txt".equals(suffix)) {
                fileList.add(file);
            }
        }
    }



    public void directoryAndFilesToList(File[] getFiles, List<File> fileList) {
        for (File file: getFiles) {
            directoryToList(file, fileList);
        }
    }

    public void indexAllDirectory(File[] getFiles, List<File> fileList) {
        for (File file: getFiles) {
            if (file.isDirectory()) {
                File[] files = file.listFiles();
                if (files != null) {
                    for (int k = 0; k < files.length; k++) {
                        if (files[k].isDirectory()) {
                            fileList.add(files[k]);
                        }
                    }
                }
            }
        }
    }
}
