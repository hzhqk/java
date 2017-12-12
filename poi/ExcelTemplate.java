
import com.storemanage.po.DeviceInfo;
import org.apache.commons.collections.CollectionUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by hqk2015@foxmail.com on 2017/11/14.
 */
public class ExcelTemplate {
    private static final Logger LOG = Logger.getLogger(ExcelTemplate.class);
    private static final int XLS_MAX_ROW = 65535; //0开始
    private static final String MAIN_SHEET_NAME = "template";
    private static final String HIDDEN_SHEET1_NAME = "hidden1";
    private static final String HIDDEN_SHEET2_NAME = "hidden2";
    private static final String WAREHOUSE_NAMES = "warehouses";
    private static final String DEVICE_NAMES = "devices";
    private static final String DEVICE_TYPE_NAMES = "deviceTypes";


    /**
     * 创建入库excel模版
     *
     * @param filePath
     * @param headers
     * @param warehouses
     * @param warehouseAndShelves
     * @return excel模版
     */
    private static File createStoreInExcelTemplate(String filePath, List<String> headers, List<String> warehouses, Map<String, List<String>> warehouseAndShelves) throws IOException {
        FileOutputStream out = null;
        File file;
        try {
            file = new File(filePath); //写文件
//            file = File.createTempFile("入库数据模版", ".xls");
            out = new FileOutputStream(file);
            HSSFWorkbook wb = new HSSFWorkbook();//创建工作薄
            HSSFSheet mainSheet = wb.createSheet(MAIN_SHEET_NAME);
            HSSFSheet dtHiddenSheet = wb.createSheet(HIDDEN_SHEET1_NAME);
            HSSFSheet wsHiddenSheet = wb.createSheet(HIDDEN_SHEET2_NAME);
            wb.setSheetHidden(1, true); //将第二个用于存储下拉框数据的sheet隐藏
            wb.setSheetHidden(2, true);
            initHeaders(wb, mainSheet, headers);
            //************* mock data ************
            DeviceInfo device1 = new DeviceInfo(1, "设备1", "01", "type1");
            DeviceInfo device2 = new DeviceInfo(2, "设备2", "02", "type2");
            DeviceInfo device3 = new DeviceInfo(3, "设备3", "03", "type3");
            List<String> deviceTypes = Arrays.asList("type1", "type2", "type3", "type4");
            List<DeviceInfo> devices = Arrays.asList(device1, device2, device3);
            initDevicesAndType(wb, dtHiddenSheet, devices, deviceTypes);
            initWarehousesAndShelves(wb, wsHiddenSheet, warehouses, warehouseAndShelves);
            initSheetNameMapping(mainSheet);
            initOtherConstraints(mainSheet);
            out.flush();
            wb.write(out);
//            file.deleteOnExit();
        } catch (Exception e) {
            e.printStackTrace();
            LOG.error("生成入库excel模版失败！", e);
            throw e;
        } finally {
            if (out != null)
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
        }
        return file;
    }

    /**
     * 初始化仓库&货架下拉框数据
     *
     * @param workbook
     * @param wsSheet
     * @param warehouses
     * @param warehousesAndShelves
     */
    private static void initWarehousesAndShelves(HSSFWorkbook workbook, HSSFSheet wsSheet, List<String> warehouses, Map<String, List<String>> warehousesAndShelves) {
        writeWarehouses(workbook, wsSheet, warehouses);
        writeShelves(workbook, wsSheet, warehouses, warehousesAndShelves);
        initWarehouseNameMapping(workbook, wsSheet.getSheetName(), warehouses.size());
    }

    /**
     * 在隐藏sheet中写入货架下拉框数据
     *
     * @param workbook
     * @param wsSheet
     * @param warehouses
     * @param warehousesAndShelves
     */
    private static void writeShelves(HSSFWorkbook workbook, HSSFSheet wsSheet, List<String> warehouses, Map<String, List<String>> warehousesAndShelves) {
        for (int i = 0; i < warehouses.size(); i++) {
            int referColNum = i + 1;
            String warehouseName = warehouses.get(i);
            List<String> shelves = warehousesAndShelves.get(warehouseName);
            if (CollectionUtils.isNotEmpty(shelves)) {
                int rowCount = wsSheet.getLastRowNum();
                if(rowCount == 0 && wsSheet.getRow(0) == null )
                    wsSheet.createRow(0);
                for (int j = 0; j < shelves.size(); j++) {
                    if (j <= rowCount) { //前面创建过的行，直接获取行，创建列
                        wsSheet.getRow(j).createCell(referColNum).setCellValue(shelves.get(j)); //设置对应单元格的值
                    } else { //未创建过的行，直接创建行、创建列
                        wsSheet.setColumnWidth(j, 4000); //设置每列的列宽
                        //创建行、创建列
                        wsSheet.createRow(j).createCell(referColNum).setCellValue(shelves.get(j)); //设置对应单元格的值
                    }
                }
            }
            initShelfNameMapping(workbook, wsSheet.getSheetName(), warehouseName, referColNum, shelves.size());
        }
    }

    /**
     * 货架与库房下拉选择数据关联
     *
     * @param workbook
     * @param wsSheetName
     * @param warehouseName
     * @param referColNum
     * @param shelfQuantity
     */
    private static void initShelfNameMapping(HSSFWorkbook workbook, String wsSheetName, String warehouseName, int referColNum, int shelfQuantity) {
        Name name = workbook.createName();
        // 设置货架名称
        name.setNameName(warehouseName);
        String referColName = getColumnName(referColNum);
        String formula = wsSheetName + "!$" + referColName + "$1:$" + referColName + "$" + shelfQuantity;
        name.setRefersToFormula(formula);
    }

    /**
     * 在隐藏sheet中写入库房下拉框数据
     *
     * @param workbook
     * @param wsSheet
     * @param warehouses
     */
    private static void writeWarehouses(HSSFWorkbook workbook, HSSFSheet wsSheet, List<String> warehouses) {
        for (int i = 0; i < warehouses.size(); i++) {
            HSSFRow row = wsSheet.createRow(i);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(warehouses.get(i));
        }
    }

    /**
     * 初始化库房选择“名称”
     *
     * @param workbook
     * @param wsSheetName
     * @param warehouseQuantity
     */
    private static void initWarehouseNameMapping(HSSFWorkbook workbook, String wsSheetName, int warehouseQuantity) {
        Name name = workbook.createName();
        // 设置仓库名称
        name.setNameName(WAREHOUSE_NAMES);
        name.setRefersToFormula(wsSheetName + "!$A$1:$A$" + warehouseQuantity);
    }

    /**
     * 根据数据值确定单元格位置（比如：0-A, 27-AB）
     *
     * @param index
     * @return
     */
    public static String getColumnName(int index) {
        StringBuilder s = new StringBuilder();
        while (index >= 26) {
            s.insert(0, (char) ('A' + index % 26));
            index = index / 26 - 1;
        }
        s.insert(0, (char) ('A' + index));
        return s.toString();
    }

    private static void initDevicesAndType(HSSFWorkbook wb, HSSFSheet dtHiddenSheet, List<DeviceInfo> devices, List<String> deviceTypes) {
        writeDevices(wb, dtHiddenSheet, devices);
        writeDeviceTypes(wb, dtHiddenSheet, deviceTypes);
        initDevicesNameMapping(wb, dtHiddenSheet.getSheetName(), devices.size());
        initDeviceTypesNameMapping(wb, dtHiddenSheet.getSheetName(), deviceTypes.size());
    }

    private static void initDeviceTypesNameMapping(HSSFWorkbook wb, String dtHiddenSheetName, int deviceTypeQuantity) {
        Name name = wb.createName();
        // 设置设备“名称”
        name.setNameName(DEVICE_TYPE_NAMES);
        name.setRefersToFormula(dtHiddenSheetName + "!$B$1:$B$" + deviceTypeQuantity);
    }

    private static void writeDeviceTypes(HSSFWorkbook wb, HSSFSheet dtHiddenSheet, List<String> deviceTypes) {
        int lastRow = dtHiddenSheet.getLastRowNum();
        if(lastRow == 0 && dtHiddenSheet.getRow(0) == null )
            dtHiddenSheet.createRow(0);
        if (CollectionUtils.isNotEmpty(deviceTypes))
            for (int j = 0; j < deviceTypes.size(); j++) {
                if (j <= lastRow) { //前面创建过的行，直接获取行，创建列
                    dtHiddenSheet.getRow(j).createCell(1).setCellValue(deviceTypes.get(j)); //设置对应单元格的值
                } else { //未创建过的行，直接创建行、创建列
                    dtHiddenSheet.setColumnWidth(j, 4000); //设置每列的列宽
                    //创建行、创建列
                    dtHiddenSheet.createRow(j).createCell(1).setCellValue(deviceTypes.get(j)); //设置对应单元格的值
                }
            }
    }

    private static void writeDevices(HSSFWorkbook wb, HSSFSheet dtHiddenSheet, List<DeviceInfo> devices) {
        for (int i = 0; i < devices.size(); i++) {
            HSSFRow row = dtHiddenSheet.createRow(i);
            HSSFCell cell1 = row.createCell(0);
            cell1.setCellValue(devices.get(i).getDeviceName());
        }
    }

    private static void initDevicesNameMapping(HSSFWorkbook workbook, String dtHiddenSheetName, int deviceQuantity) {
        Name name = workbook.createName();
        // 设置设备“名称”
        name.setNameName(DEVICE_NAMES);
        name.setRefersToFormula(dtHiddenSheetName + "!$A$1:$A$" + deviceQuantity);
    }

    /**
     * 初始化表头
     *
     * @param wb
     * @param mainSheet
     * @param headers
     */
    private static void initHeaders(HSSFWorkbook wb, HSSFSheet mainSheet, List<String> headers) {
        //表头样式
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 创建一个居中格式
        //字体样式
        HSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("微软雅黑");
        fontStyle.setFontHeightInPoints((short) 12);
        fontStyle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(fontStyle);
        //生成sheet1内容
        HSSFRow rowFirst = mainSheet.createRow(0);//第一个sheet的第一行为标题
        mainSheet.createFreezePane(0, 1, 0, 1); //冻结第一行
        //写标题
        for (int i = 0; i < headers.size(); i++) {
            HSSFCell cell = rowFirst.createCell(i); //获取第一行的每个单元格
            mainSheet.setColumnWidth(i, 4000); //设置每列的列宽
            cell.setCellStyle(style); //加样式
            cell.setCellValue(headers.get(i)); //往单元格里写数据
        }
    }

    /**
     * 主sheet中下拉框初始化
     *
     * @param mainSheet
     */
    private static void initSheetNameMapping(HSSFSheet mainSheet) {
        DataValidation warehouseValidation = getDataValidationByFormula(WAREHOUSE_NAMES, 3);
        DataValidation shelfValidation = getDataValidationByFormula("INDIRECT($D1)", 4); //formula同"INDIRECT(INDIRECT(\"R\"&ROW()&\"C\"&(COLUMN()-1),FALSE))"
        DataValidation deviceValidation = getDataValidationByFormula(DEVICE_NAMES, 0);
        DataValidation deviceTypeValidation = getDataValidationByFormula(DEVICE_TYPE_NAMES, 1);
        // 主sheet添加验证数据
        mainSheet.addValidationData(warehouseValidation);
        mainSheet.addValidationData(shelfValidation);
        mainSheet.addValidationData(deviceValidation);
        mainSheet.addValidationData(deviceTypeValidation);
    }

    /**
     * 初始化其他列输入限定
     *
     * @param mainSheet
     */
    private static void initOtherConstraints(HSSFSheet mainSheet) {
        DataValidation quantityValidation = getDecimalValidation(1, XLS_MAX_ROW, 2);
        DataValidation tierValidation = getDecimalValidation(1, XLS_MAX_ROW, 5);
        DataValidation colValidation = getDecimalValidation(1, XLS_MAX_ROW, 6);
        mainSheet.addValidationData(tierValidation);
        mainSheet.addValidationData(colValidation);
        mainSheet.addValidationData(quantityValidation);
    }
    /**
     * 生成下拉框及提示
     *
     * @param formulaString
     * @param columnIndex
     * @return
     */
    public static DataValidation getDataValidationByFormula(String formulaString, int columnIndex) {
        // 加载下拉列表内容
        DVConstraint constraint = DVConstraint.createFormulaListConstraint(formulaString);
        // 设置数据有效性加载在哪个单元格上。
        // 四个参数分别是：起始行、终止行、起始列、终止列
        CellRangeAddressList regions = new CellRangeAddressList(1, XLS_MAX_ROW, columnIndex, columnIndex);
        // 数据有效性对象
        DataValidation dataValidationList = new HSSFDataValidation(regions, constraint);
        dataValidationList.createErrorBox("Error", "请选择或输入有效的选项，或下载最新模版重试！");
        String promptText = initPromptText(columnIndex);
        dataValidationList.createPromptBox("", promptText);
        return dataValidationList;
    }

    /**
     * 数字输入验证
     * @param firstRow
     * @param lastRow
     * @param columnIndex
     * @return
     */
    private static DataValidation getDecimalValidation(int firstRow, int lastRow, int columnIndex) {
        // 创建一个规则：>0的整数
        DVConstraint constraint = DVConstraint.createNumericConstraint(DVConstraint.ValidationType.INTEGER, DVConstraint.OperatorType.GREATER_OR_EQUAL, "0", "0");
        // 设定在哪个单元格生效
        CellRangeAddressList regions = new CellRangeAddressList(firstRow, lastRow, columnIndex, columnIndex);
        // 创建规则对象
        HSSFDataValidation decimalVal = new HSSFDataValidation(regions, constraint);
        decimalVal.createPromptBox("", initPromptText(columnIndex));
        decimalVal.createErrorBox("输入值类型或大小有误！", "数值型，请输入大于0 的整数。");
        return decimalVal;
    }

    /**
     * 初始化下拉框提示信息
     *
     * @param columnIndex
     * @return
     */
    private static String initPromptText(int columnIndex) {
        String promptText ="";
        //custom column prompt
        switch (columnIndex) {
            case 2:
                promptText = "请输入大于0的整数！";
                break;
            case 4:
                promptText = "请下拉选择或输入有效项！且先选择库房！";
                break;
            case 5:
                promptText = "请输入大于0的整数！";
                break;
            case 6:
                promptText = "请输入大于0的整数！";
                break;
        }
        return promptText;
    }

    public static File test() throws IOException {
        List<String> headers = Arrays.asList("设备名称", "设类型", "数量", "存放库房", "存放货架", "存放层", "存放列");
        List<String> warehouses = Arrays.asList("库房1", "库房2");
        Map<String, List<String>> warehousesAndShelves = new HashMap<>();
        warehousesAndShelves.put("库房1", Arrays.asList("货架1-1", "货架1-2", "货架1-3"));
        warehousesAndShelves.put("库房2", Arrays.asList("货架2-1", "货架2-2", "货架2-3"));
        return ExcelTemplate.createStoreInExcelTemplate("f:/test.xls", headers, warehouses, warehousesAndShelves);
    }

    public static void main(String[] args) throws IOException {
        ExcelTemplate.test();
    }
}
