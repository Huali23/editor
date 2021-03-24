import bean.DeviceBoxItem;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.FocusEvent;
import java.awt.event.FocusListener;
import java.io.*;
import java.util.*;
import java.util.List;

public class MyFrame extends JFrame {

    private static final String FILE_DIC = "C:\\temp\\editor"; //windows系统资源目录
    private static final String FILE_CODE_DIC = "C:\\temp\\editor\\code"; //windows系统资源目录
    private static final String FILE_CODE_NAME = "file2Code.xlsx";

    private JTextField fileField = new JTextField(16); //文件名 16代表宽度 手动输入
    private JTextField projectField = new JTextField(16); //项目代码
    private JComboBox<String> systemBox = new JComboBox<>(); //系统代码
    private JComboBox<String> activeBox = new JComboBox<>(); //主要活动代码
    private JComboBox<String> stageBox = new JComboBox<>(); //阶段代码
    private JComboBox<String> deviceBox = new JComboBox<>(); //设备代码
    private JComboBox<String> subDeviceBox = new JComboBox<>(); //分享设备代码
    private JComboBox<String> fileBox = new JComboBox<>(); //文件类型
    private JComboBox<String> editorBox = new JComboBox<>(); //编发者所属部门
    private JLabel numberLabel = new JLabel(); //文件顺序号

    private JButton button = new JButton("确定");

    private static Map<String, String> fileNameMap = new HashMap<>();
    private static Map<String, String> systemMap = new HashMap<>();
    private static Map<String, String> activeMap = new HashMap<>();
    private static Map<String, String> stageMap = new HashMap<>();
    private static Map<String, String> fileMap = new HashMap<>();
    private static Map<String, String> editorMap = new HashMap<>();
    private static List<DeviceBoxItem> deviceList = new ArrayList<>();

    public MyFrame(String title) {
        super(title);
        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(11, 3, 10, 15));
        this.mapFactory();
        this.BoxFactory();
        this.frameFactory(panel);
        this.add(panel);
        this.bindListener();
    }

    private void bindListener() {
        deviceBox.addItemListener(e -> {
            onDeviceItemChange();
        });
        button.addActionListener((e) -> {
            onButtonClick();
        });
        fileField.addFocusListener(new FocusListener() {
            @Override
            public void focusGained(FocusEvent e) {

            }

            @Override
            public void focusLost(FocusEvent e) {

            }
        });
    }


    private void onDeviceItemChange() {
        String text = (String) deviceBox.getSelectedItem();
        this.addSubDeviceItem(text);
    }


    private void frameFactory(JPanel panel) {
        panel.add(new JLabel("文件名"));
        panel.add(fileField);
        panel.add(new JLabel("项目代码"));
        panel.add(projectField);
        panel.add(new JLabel("系统代码"));
        panel.add(systemBox);
        panel.add(new JLabel("主要活动代码"));
        panel.add(activeBox);
        panel.add(new JLabel("阶段代码"));
        panel.add(stageBox);
        panel.add(new JLabel("设备代码"));
        panel.add(deviceBox);
        panel.add(new JLabel("分项设备代码"));
        panel.add(subDeviceBox);
        panel.add(new JLabel("文件类型"));
        panel.add(fileBox);
        panel.add(new JLabel("编发者所属部门"));
        panel.add(editorBox);
        panel.add(new JLabel("文件顺序号"));
        panel.add(numberLabel);
        panel.add(new JLabel());
        panel.add(button);
    }

    private void onButtonClick() {
        if (StringUtils.isBlank(fileField.getText()) || StringUtils.isBlank(projectField.getText())) {
            JOptionPane.showMessageDialog(this, "输入框为空");
            return;
        } else if (fileNameMap.containsKey(fileField.getText())) {
            JOptionPane.showMessageDialog(this, "文件名已存在");
            return;
        }
        StringBuilder result = new StringBuilder();
        String fileName = fileField.getText();
        result.append(projectField.getText());
        result.append('-').append(systemMap.get(systemBox.getSelectedItem()));
        result.append('-').append(activeMap.get(activeBox.getSelectedItem()));
        result.append('-').append(stageMap.get(stageBox.getSelectedItem()));
        result.append('-').append(this.getDeviceValue(deviceBox.getSelectedItem()));
        result.append('-').append(this.getDeviceValue(subDeviceBox.getSelectedItem()));
        result.append('-').append(fileMap.get(fileBox.getSelectedItem()));
        result.append('-').append(editorMap.get(editorBox.getSelectedItem()));
        result.append('-').append(this.getNumberCode(result.toString()));
        fileNameMap.put(fileName, result.toString());
        this.saveFileCode(fileName, result.toString());
        JOptionPane.showMessageDialog(this, "转换完成");
    }

    private void saveFileCode(String fileName, String result) {
        InputStream inputStream = null;
        FileOutputStream outputStream = null;
        try{
            inputStream = new FileInputStream(FILE_DIC + File.separator + FILE_CODE_NAME);
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFRow row = sheet.createRow(sheet.getPhysicalNumberOfRows());
            row.createCell(0).setCellValue(fileName);
            row.createCell(1).setCellValue(result);
            outputStream = new FileOutputStream(FILE_DIC + File.separator + FILE_CODE_NAME);
            workbook.write(outputStream);
            outputStream.close();
        } catch (FileNotFoundException e) {
            System.out.println("excel输出失败");
        } catch (IOException e) {
            System.out.println("IO异常");
        } finally {
            try {
                assert inputStream != null;
                inputStream.close();
                assert outputStream != null;
                outputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private StringBuilder getNumberCode(String result) {
        int i = 1;
        for (String str : fileNameMap.values()) {
            if (str.contains(result)) {
                i++;
            }
        }
        StringBuilder numberCode = new StringBuilder(String.valueOf(i));
        for (int j = numberCode.length(); j < 3; j++) {
            numberCode.insert(0, "0");
        }
        return numberCode;
    }

    private String getDeviceValue(Object selectedItem) {
        if (StringUtils.isBlank((CharSequence) selectedItem)) {
            return "X";
        } else {
            for (DeviceBoxItem deviceBoxItem : deviceList) {
                if (selectedItem.equals(deviceBoxItem.getName())) {
                    return deviceBoxItem.getValue();
                }
            }
            return "null";
        }
    }


    /**
     * subDeviceBox 更新下拉
     *
     * @param name name
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void addSubDeviceItem(String name) {
        if (StringUtils.isNotBlank(name)) {
            subDeviceBox.removeAllItems();
            for (DeviceBoxItem deviceBoxItem : deviceList) {
                if (name.equals(deviceBoxItem.getParentName())) {
                    subDeviceBox.addItem(deviceBoxItem.getName());
                }
            }
        }

    }

    /**
     * map初始化
     *
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void mapFactory() {
//        this.excel2Map(fileNameMap, FILE_DIC + File.separator + FILE_CODE_NAME);
        this.excel2Map(systemMap, FILE_CODE_DIC + File.separator + "systemCode.xlsx");
        this.excel2Map(activeMap, FILE_CODE_DIC + File.separator + "activeCode.xlsx");
        this.excel2Map(stageMap, FILE_CODE_DIC + File.separator + "stageCode.xlsx");
        this.excel2Map(fileMap, FILE_CODE_DIC + File.separator + "fileCode.xlsx");
        this.excel2Map(editorMap, FILE_CODE_DIC + File.separator + "editorCode.xlsx");
        this.excel2DeviceList(deviceList, FILE_CODE_DIC + File.separator + "deviceCode.xlsx");
    }

    /**
     * 下拉框初始化
     *
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void BoxFactory() {
        this.map2BoxItem(systemMap, systemBox);
        this.map2BoxItem(activeMap, activeBox);
        this.map2BoxItem(stageMap, stageBox);
        this.map2BoxItem(fileMap, fileBox);
        this.map2BoxItem(editorMap, editorBox);
        this.list2DeviceBoxItem(deviceList, deviceBox);
        if (deviceList.size() > 0) {
            this.addSubDeviceItem(deviceList.get(0).getName());
        }
    }

    /**
     * device list 转 下拉框下拉选项
     *
     * @param deviceList list
     * @param deviceBox  下拉框
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void list2DeviceBoxItem(List<DeviceBoxItem> deviceList, JComboBox<String> deviceBox) {
        for (DeviceBoxItem deviceBoxItem : deviceList) {
            if (StringUtils.isBlank(deviceBoxItem.getParentName())) {
                deviceBox.addItem(deviceBoxItem.getName());
            }
        }
    }

    /**
     * 将map添加到下拉框下拉选项
     *
     * @param map      map
     * @param comboBox 下拉框
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void map2BoxItem(Map<String, String> map, JComboBox<String> comboBox) {
        for (String str : map.keySet()) {
            comboBox.addItem(str);
        }
    }

    /**
     * excel转map
     *
     * @param filePath 文件绝对路径
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void excel2Map(Map<String, String> map, String filePath) {
        for (Row row : this.getSheet(filePath)) {
            if (row.getRowNum() == 0) {
                continue;
            }
            map.put(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue());
        }
    }

    /**
     * device excel 转 deviceList
     * API描述
     *
     * @param deviceList list
     * @param filePath   excel绝对路径
     * @return void
     * @Author liujingeng
     * @Date 2020/2/27
     */
    private void excel2DeviceList(List<DeviceBoxItem> deviceList, String filePath) {
        DeviceBoxItem deviceBoxItem;
        for (Row row : this.getSheet(filePath)) {
            if (row.getRowNum() == 0) {
                continue;
            }
            deviceBoxItem = new DeviceBoxItem(row.getPhysicalNumberOfCells() == 2 ? "" : row.getCell(0).getStringCellValue(),
                    row.getCell(1).getStringCellValue(),
                    row.getCell(2).getStringCellValue());
            deviceList.add(deviceBoxItem);
        }
    }

    private XSSFSheet getSheet(String filePath) {
        XSSFSheet sheet = null;
        try (
                InputStream inputStream = new FileInputStream(filePath);
        ) {
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            sheet = workbook.getSheetAt(0);

        } catch (FileNotFoundException e) {
            System.out.println("文件获取异常");
        } catch (IOException e) {
            System.out.println("文件转换异常");
        }
        return sheet;
    }

}
