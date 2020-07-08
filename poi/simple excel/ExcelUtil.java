package excel;

import com.google.common.collect.Lists;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 简单excel生成工具(使用POI)，配合@PoiExcelFiled注解使用
 * <p>
 * 返回HSSFSheet可以在此基础上新增sheet、获取sheet字节数组
 * </p>
 *
 * @author hzhqk
 * @date 2018/03/21
 */
@Slf4j
public class ExcelUtil {
    /**
     * 默认sheet名称
     */
    private static final String DEFAULT_SHEET_NAME = "sheet";
    /**
     * 单个Sheet页最大行数（除去标题）
     */
    private static final int SINGLE_SHEET_MAX_ROWS = 65535;
    /**
     * 默认单元格宽度
     */
    private static final int DEFAULT_CELL_WIDTH = 5000;
    /**
     * 默认日期格式
     */
    private static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";

    /**
     * 根据列标题和列数据生成excel表格文件
     * excel默认一个sheet且sheet名称默认为sheet1
     * <pre>
     *     获取excel bytes示例：
     *         ByteArrayOutputStream os = new ByteArrayOutputStream();
     *         workbook.write(os);
     *         或直接使用 PoiExcelUtil.getExcelBytes(hssfWorkbook);
     * </pre>
     *
     * @param datas
     * @return 没有使用write方法时要手动关闭HSSFWorkbook
     * @throws Exception
     */
    public static HSSFWorkbook createExcel(List<?> datas) throws Exception {
        return createExcelWithSheetName(null, datas);
    }

    /**
     * 根据列标题和列数据生成excel表格文件，excel文件默认一个sheet
     *
     * @param sheetName 生成的excel文件的sheet名，分页时会在后面加序号
     * @param datas     列数据
     * @return 文件byte
     * @throws Exception
     */
    public static HSSFWorkbook createExcelWithSheetName(String sheetName, List<?> datas) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        createSheetAndWriteData(workbook, sheetName, datas);
        return workbook;
    }

    /**
     * 在现有HSSFWorkbook基础上新建sheet页
     *
     * @param workbook
     * @param sheetName
     * @param datas
     * @return
     * @throws Exception
     */
    public static HSSFWorkbook createSheetAndWriteData(HSSFWorkbook workbook, String sheetName, List<?> datas) throws Exception {
        if (StringUtils.isBlank(sheetName)) {
            sheetName = getDefaultSheetName(workbook, workbook.getNumberOfSheets());
        }
        if (CollectionUtils.isEmpty(datas)) {
            log.info("数据为空，不写入数据");
            // 创建一个空sheet，防止打开报错
            workbook.createSheet(sheetName);
            return workbook;
        }
        List<Field> allFields = getAllFields(datas.get(0).getClass());
        List<Field> dataFields = allFields.stream().filter(f -> f.getAnnotation(ExcelField.class) != null).collect(Collectors.toList());
        if (datas.size() > SINGLE_SHEET_MAX_ROWS) {
            List<? extends List<?>> parts = Lists.partition(datas, SINGLE_SHEET_MAX_ROWS);
            int sheetIndex = 1;
            for (List<?> part : parts) {
                createOneSheetAndWriteData(workbook, sheetName + sheetIndex++, part, dataFields);
            }
        } else {
            createOneSheetAndWriteData(workbook, sheetName, datas, dataFields);
        }
        return workbook;
    }

    private static String getDefaultSheetName(HSSFWorkbook workbook, int sheetNum) {
        if (sheetNum == 0) {
            return DEFAULT_DATE_PATTERN + 1;
        }
        String sheetName = DEFAULT_SHEET_NAME + (sheetNum + 1);
        if (workbook.getSheetIndex(sheetName) >= 0) {
            sheetName = getDefaultSheetName(workbook, sheetNum + 1);
        }
        return sheetName;
    }

    private static void createOneSheetAndWriteData(HSSFWorkbook workbook, String sheetName, List<?> datas, List<Field> dataFields) throws Exception {
        HSSFSheet sheet = workbook.createSheet(sheetName);
        List<FieldWithFormatter> fieldWithFormatter = convert2FieldWithFormatter(dataFields);
        initSheetHeaders(workbook, sheet, fieldWithFormatter);
        writeData(sheet, datas, fieldWithFormatter);
    }

    /**
     * 将HSSFWorkbook转为bytes
     *
     * @param workbook
     * @return
     * @throws IOException
     */
    public static byte[] getExcelBytes(HSSFWorkbook workbook) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        return os.toByteArray();
    }

    /**
     * 传到前台
     *
     * @param workbook
     * @param response
     */
    public static void write2Response(HSSFWorkbook workbook, String fileName, HttpServletResponse response) throws IOException {
        response.setHeader("Content-Disposition", "attachment;filename=" + new String((fileName).getBytes(), "ISO-8859-1"));
        response.setContentType("application/x-execl");
        workbook.write(response.getOutputStream());
    }

    private static void writeData(HSSFSheet sheet, List<?> datas, List<FieldWithFormatter> fieldWithFormatter) throws Exception {
        if (CollectionUtils.isEmpty(datas)) {
            return;
        }
        int startX = 1;
        int startY = 0;
        for (Object item : datas) {
            HSSFRow row = sheet.createRow(startX);
            for (FieldWithFormatter fwf : fieldWithFormatter) {
                HSSFCell cell = row.createCell(startY++);
                Object origin = fwf.getField().get(item);
                String columnData = "";
                if (null != origin) {
                    ExcelColumnFormatter formatter = fwf.getFormatter();
                    if (formatter != null) {
                        columnData = formatter.format(origin);
                    } else {
                        columnData = origin.toString();
                    }
                }
                cell.setCellValue(columnData);
            }
            startX++;
            startY = 0;
        }
    }

    private static List<FieldWithFormatter> convert2FieldWithFormatter(List<Field> fields) {
        return fields.stream().map(f -> {
            ExcelField annotation = f.getAnnotation(ExcelField.class);
            ExcelColumnFormatter formatter = null;
            try {
                Class<? extends ExcelColumnFormatter> format = annotation.formatter();
                if (format != NoFormatter.class) {
                    formatter = format.newInstance();
                }
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
            return new FieldWithFormatter(f, annotation.name(), annotation.order(), formatter);
        }).sorted(Comparator.comparingInt(f -> f.order)).collect(Collectors.toList());
    }

    private static class FieldWithFormatter {
        private Field field;
        private String columnName;
        private int order;
        private ExcelColumnFormatter formatter;

        public FieldWithFormatter(Field field, String columnName, int order, ExcelColumnFormatter formatter) {
            this.field = field;
            this.columnName = columnName;
            this.order = order;
            this.formatter = formatter;
        }

        public Field getField() {
            return field;
        }

        public void setField(Field field) {
            this.field = field;
        }

        public String getColumnName() {
            return columnName;
        }

        public void setColumnName(String columnName) {
            this.columnName = columnName;
        }

        public int getOrder() {
            return order;
        }

        public void setOrder(int order) {
            this.order = order;
        }

        public ExcelColumnFormatter getFormatter() {
            return formatter;
        }

        public void setFormatter(ExcelColumnFormatter formatter) {
            this.formatter = formatter;
        }
    }

    private static <T> List<Field> getAllFields(Class<T> clz) {
        List<Field> fields = Lists.newArrayList();
        Class<?> tmpClz = clz;
        // 不获取Object层的属性
        String finalParent = "java.lang.object";
        while (tmpClz != null && !tmpClz.getName().toLowerCase().equals(finalParent)) {
            // 只获取bean普通属性
            for (Field field : tmpClz.getDeclaredFields()) {
                // 不在设置数据时设置访问权限
                field.setAccessible(true);
                int modifiers = field.getModifiers();
                if (modifiers == Modifier.PUBLIC || modifiers == Modifier.PRIVATE || modifiers == Modifier.PROTECTED) {
                    fields.add(field);
                }
            }
            tmpClz = tmpClz.getSuperclass();
        }
        return fields;
    }

    /**
     * 初始化表头
     *  @param wb
     * @param sheet
     * @param dataFields
     */
    private static void initSheetHeaders(HSSFWorkbook wb, HSSFSheet sheet, List<FieldWithFormatter> dataFields) {
        if (CollectionUtils.isEmpty(dataFields)) {
            return;
        }
        // 表头样式
        HSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
        style.setAlignment(HorizontalAlignment.CENTER);
        // 字体样式
        HSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("微软雅黑");
        fontStyle.setFontHeightInPoints((short) 12);
        fontStyle.setBold(true);
        style.setFont(fontStyle);
        // 生成sheet1内容
        // 第一个sheet的第一行为标题
        HSSFRow rowFirst = sheet.createRow(0);
        // 冻结第一行
        sheet.createFreezePane(0, 1, 0, 1);
        // 写标题
        for (int i = 0; i < dataFields.size(); i++) {
            // 获取第一行的每个单元格
            HSSFCell cell = rowFirst.createCell(i);
            // 设置每列的列宽
            sheet.setColumnWidth(i, DEFAULT_CELL_WIDTH);
            //加样式
            cell.setCellStyle(style);
            //往单元格里写数据
            cell.setCellValue(dataFields.get(i).getColumnName());
        }
    }

    /**
     * test case 运行前临时注释掉servlet-api包的scope，测试完放开
     *
     * @param args
     * @throws Exception
     */
    public static void main(String[] args) throws Exception {
        String str = "字符串";
        Date date = new Date();
        int mint = 12345;
        float mfloat = 1.2345f;

        List<Date> list = Arrays.asList(date, date, date);
        Map<String, Date> map = new HashMap<String, Date>();
        map.put("date1", date);
        map.put("data2", date);

        TestData mb = new TestData(str, date, mint, mfloat, new BigDecimal("0.9999"), list, map);
        List<TestData> datas = Lists.newArrayList();
        for (int i = 0; i < 70000; i++) {
            datas.add(mb);
        }
        HSSFWorkbook wb = ExcelUtil.createExcelWithSheetName("测试", datas);
        ExcelUtil.createSheetAndWriteData(wb, "", Arrays.asList(new Extension(9999)));
        wb.write(new FileOutputStream(new File("/Users/shhanqiankun/Desktop/excel.xls")));
        System.out.println("everything goes well");
    }

    @AllArgsConstructor
    @Data
    private static class Extension {
        @ExcelField(name = "column")
        private Integer total;
    }

    @AllArgsConstructor
    @Data
    private static class TestData {
        @ExcelField(name = "String", order = 9)
        private String str;
        @ExcelField(name = "Date", order = 8, formatter = DefaultDateFormatter.class)
        private Date date;
        @ExcelField(name = "Integer", order = 7)
        private int mint;
        @ExcelField(name = "Float", order = 6)
        private float mfloat;
        @ExcelField(name = "Decimal", order = 3)
        private BigDecimal bigDecimal;
        @ExcelField(name = "list", order = 999)
        private List<Date> list;
        @ExcelField(name = "Map")
        private Map<String, Date> map;
    }
}
