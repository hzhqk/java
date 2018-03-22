package com.storemanage.util;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 简单excel生成工具(使用POI)
 * Created by HanQiankun(hqk2015@foxmail.com) on 2018/03/22.
 */
public class SimpleExcelUtil {
    /** 默认sheet名称 */
    private static final String DEFAULT_SHEET_NAME = "sheet";
    /** 单个Sheet页最大行数 */
    private static final int SINGLE_SHEET_MAX_ROWS = 65535;
    /** 默认单元格宽度 */
    private static final int DEFAULT_CELL_WIDTH = 5000;
    /** 默认日期格式 */
    private static final String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";
    /** 默认日期格式 */
    private static final ThreadLocal<SimpleDateFormat> DEFAULT_DATE_FORMAT = new ThreadLocal<SimpleDateFormat>(){
        @Override
        protected SimpleDateFormat initialValue() {
            return new SimpleDateFormat(DEFAULT_DATE_PATTERN);
        }
    };

    /**
     * 根据列标题和列数据生成excel表格文件
     * excel默认一个sheet且sheet名称默认为sheet1
     * 日期格式：yyyy-MM-dd HH:mm:ss
     * list格式：XXX,XXXX,XXXX
     * map格式：a:1,b:2
     *
     * @param columnTitles
     *            列标题
     * @param datas
     *            列数据，只支持对象的属性使用string、date、Number数字类型和list、map，有需要请扩展
     * @param columnFields
     *            列标题对应对象的属性名（依次与列标题对应）
     * @return
     * @throws Exception
     */
    public static byte[] createExcel(List<String> columnTitles, List<? extends Object> datas, List<String> columnFields) throws Exception {
        return createExcel(null, columnTitles, datas, columnFields,null);
    }

    /**
     * 根据列标题和列数据生成excel表格文件，生成excel默认一个sheet且sheet名称默认为sheet1
     * 日期格式：yyyy-MM-dd HH:mm:ss
     * list格式：XXX,XXXX,XXXX
     * map格式：a:1,b:2
     *
     * @param columnTitles
     *            列标题
     * @param datas
     *            列数据，只支持对象的属性使用string、date、Number数字类型和list、map，有需要请扩展
     * @param columnFields
     *            列标题对应对象的属性名（依次与列标题对应）
     * @param customFieldFormats
     * 			  每列数据自定义format,没有为空，list类型属性:格式化其子元素，map类型属性: 格式化其子元素的value
     * @return
     * @throws Exception
     */
    public static byte[] createExcel( List<String> columnTitles, List<? extends Object> datas, List<String> columnFields,
                                     Map<String,? extends Format> customFieldFormats) throws Exception {
        return createExcel(null, columnTitles, datas, columnFields, customFieldFormats);
    }

    /**
     * 根据列标题和列数据生成excel表格文件
     * excel默认一个sheet且sheet名称默认为sheet1
     * 日期格式：yyyy-MM-dd HH:mm:ss
     * list格式：XXX,XXXX,XXXX
     * map格式：a:1,b:2
     *
     * @param fileName
     *              文件名称
     * @param columnTitles
     *            列标题
     * @param datas
     *            列数据，只支持对象的属性使用string、date、Number数字类型和list、map，有需要请扩展
     * @param columnFields
     *            列标题对应对象的属性名（依次与列标题对应）
     * @return
     * @throws Exception
     */
    public static File createExcel(String fileName, List<String> columnTitles, List<? extends Object> datas, List<String> columnFields) throws Exception {
        byte[] fileBytes = createExcel(null, columnTitles, datas, columnFields,null);
        File file = createTmpFile(fileName);
        FileUtils.writeByteArrayToFile(file, fileBytes);
        return file;
    }

    /**
     * 根据列标题和列数据生成excel表格文件，excel文件默认一个sheet
     * 日期格式：yyyy-MM-dd HH:mm:ss
     * list格式：XXX,XXXX,XXXX
     * map格式：a:1,b:2
     *
     * @param sheetName
     *            生成的excel文件的sheet名
     * @param columnTitles
     *            列标题
     * @param datas
     *            列数据，只支持对象的属性使用string、date、Number数字类型和list、map，有需要请扩展
     * @param columnFields
     *            列标题对应对象的属性名（依次与列标题对应）
     * @param customFieldFormats
     * 			    对象属性自定义format,没有为空，list类型属性:格式化其子元素，map类型属性: 格式化其子元素的value
     * @return 文件byte
     * @throws Exception
     */
    public static byte[] createExcel(String sheetName, List<String> columnTitles, List<? extends Object> datas,
                                   List<String> columnFields, Map<String,? extends Format> customFieldFormats) throws Exception {
        if (datas == null || columnTitles == null) {
            throw new IllegalArgumentException("illegal data: data or columnTitles is null");
        }
        if (columnTitles.size() != columnFields.size()) {
            throw new IllegalArgumentException("every column title should have its mapped column field name");
        }
//        File file = null;
        HSSFWorkbook workbook = new HSSFWorkbook();
        int dataSize = datas.size();
        int part = dataSize / SINGLE_SHEET_MAX_ROWS;
        if(dataSize > SINGLE_SHEET_MAX_ROWS) {
            sheetName = StringUtils.isNotBlank(sheetName) ? sheetName : DEFAULT_SHEET_NAME;
            for(int i = 0; i < part; i++) {
                String tmpSheetName = String.format("%s%d", sheetName, i + 1);
                createSheetAndWriteData(tmpSheetName, columnTitles, datas.subList(0, SINGLE_SHEET_MAX_ROWS), columnFields,
                        customFieldFormats, workbook);
                datas.subList(0, SINGLE_SHEET_MAX_ROWS).clear();
            }
            if(CollectionUtils.isNotEmpty(datas)) {
                sheetName = String.format("%s%d", sheetName, part + 1);
                createSheetAndWriteData(sheetName, columnTitles, datas, columnFields, customFieldFormats, workbook);
            }
        } else {
            sheetName = StringUtils.isNotBlank(sheetName) ? sheetName : DEFAULT_SHEET_NAME + "1";
            createSheetAndWriteData(sheetName, columnTitles, datas, columnFields,customFieldFormats, workbook);
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        return os.toByteArray();
    }

    private static void createSheetAndWriteData(String sheetName, List<String> columnTitles, List<? extends Object> datas,
                                                List<String> columnFields, Map<String,? extends Format> customFieldFormats,
                                                HSSFWorkbook workbook) throws Exception {
        HSSFSheet sheet = workbook.createSheet(sheetName);
        initSheetHeaders(workbook, sheet, columnTitles);
        writeData(sheet, datas, columnFields,customFieldFormats);
    }

    private static void writeData(HSSFSheet sheet, List<? extends Object> datas, List<String> columnFields,
                                  Map<String,? extends Format> customFieldFormats) throws Exception {
        int startX = 1;
        int startY = 0;
        for (Object item : datas) {
            HSSFRow row = sheet.createRow(startX);
            for (int i = 0; i < columnFields.size(); i++) {
                String fieldName = columnFields.get(i);
                Field field = item.getClass().getDeclaredField(fieldName);
                field.setAccessible(true);
                Format fmt = null;
                if(customFieldFormats != null && customFieldFormats.containsKey(fieldName)) {
                    fmt = customFieldFormats.get(fieldName);
                }
                HSSFCell cell = row.createCell(startY++);
                String columnData = resolveFieldData(field.get(item),fmt);
                cell.setCellValue(columnData);
            }
            startX++;
            startY = 0;
        }
    }

    @SuppressWarnings("unchecked")
    private static String resolveFieldData(Object orgnColumnData,Format fmt) {
        String columnData = "";
        if (orgnColumnData == null) {
        } else if (orgnColumnData instanceof String) {
            columnData = fmt == null ? (String) orgnColumnData : fmt.format((String) orgnColumnData);
        } else if (orgnColumnData instanceof Date) {
            columnData = fmt ==null ? formatDate((Date) orgnColumnData) : fmt.format((Date) orgnColumnData);
        } else if (orgnColumnData instanceof Number) {
            columnData = fmt == null ? String.valueOf(orgnColumnData) : fmt.format((Number)orgnColumnData);
        } else if (orgnColumnData.getClass().isArray()) {
            List<Object> toList = Arrays.asList((Object[]) orgnColumnData);
            columnData = resolveList(toList,fmt);
        } else if (orgnColumnData instanceof List) {
            List<Object> toList = (List<Object>) orgnColumnData;
            columnData = resolveList(toList,fmt);
        } else if (orgnColumnData instanceof Map) {
            columnData = resolveMap((Map<Object, Object>) orgnColumnData,fmt);
        } else {
            throw new IllegalArgumentException("unsupported field type");
        }
        return columnData;
    }

    private static String resolveList(List<Object> list,Format fmt) {
        StringBuilder builder = new StringBuilder();
        for (Object obj : list) {
            builder.append(resolveFieldData(obj, fmt)).append(",");
        }
        if(builder.length() > 0) {
            return builder.substring(0, builder.length() - 1);
        }
        return builder.toString();
    }

    private static String resolveMap(Map<Object, Object> map,Format fmt) {
        StringBuilder builder = new StringBuilder();
        Iterator<Map.Entry<Object, Object>> iterator = map.entrySet().iterator();
        while(iterator.hasNext()) {
            Map.Entry<Object, Object> entry = iterator.next();
            String mKey = entry.getKey() != null ? entry.getKey().toString() : "";
            builder.append(mKey).append(":").append(resolveFieldData(entry.getValue(),fmt)).append(",");
        }
        if(builder.length() > 0) {
            return builder.substring(0, builder.length() - 1);
        }
        return builder.toString();
    }

    /**
     * 初始化表头
     *
     * @param wb
     * @param sheet
     * @param headers
     */
    private static void initSheetHeaders(HSSFWorkbook wb, HSSFSheet sheet, List<String> headers) {
        // 表头样式
        HSSFCellStyle style = wb.createCellStyle();
        // 创建一个居中格式
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        // 字体样式
        HSSFFont fontStyle = wb.createFont();
        fontStyle.setFontName("微软雅黑");
        fontStyle.setFontHeightInPoints((short) 12);
        fontStyle.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(fontStyle);
        // 生成sheet1内容
        // 第一个sheet的第一行为标题
        HSSFRow rowFirst = sheet.createRow(0);
        // 冻结第一行
        sheet.createFreezePane(0, 1, 0, 1);
        // 写标题
        for (int i = 0; i < headers.size(); i++) {
            // 获取第一行的每个单元格
            HSSFCell cell = rowFirst.createCell(i);
            // 设置每列的列宽
            sheet.setColumnWidth(i, DEFAULT_CELL_WIDTH);
            //加样式
            cell.setCellStyle(style);
            //往单元格里写数据
            cell.setCellValue(headers.get(i));
        }
    }

    private static String formatDate(Date date) {
        return DEFAULT_DATE_FORMAT.get().format(date);
    }

    private static File createTmpFile(String fileName) throws IOException {
        String tmpFileName = StringUtils.isNotBlank(fileName) ? fileName : new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        File file = File.createTempFile(tmpFileName + "_", ".xls");
        file.deleteOnExit();
        return file;
    }

    // test case
    public static void main(String[] args) throws Exception {
        String str = "字符串";
        Date date = new Date();
        int mint = 12345;
        float mfloat = 1.2345f;
        List<Date> list = Arrays.asList(date,date,date);
        Map<String,Date> map = new HashMap<String, Date>();
        map.put("date1", date);
        map.put("data2", date);

        MyJxlTestBean mb = new MyJxlTestBean(str, date, mint, mfloat, list, map);
        List<MyJxlTestBean> datas = new ArrayList<MyJxlTestBean>(Arrays.asList(mb,mb,mb));
        List<String> columnTitles = Arrays.asList("字符串", "日期", "整数", "小数","list","map");
        List<String> columnFields = new ArrayList<String>(Arrays.asList("str", "date", "mint", "mfloat","list","map"));
        Map<String,Format> fmts = new HashMap<String, Format>();
        DateFormat df = new SimpleDateFormat("yyyy*MM*dd");
        fmts.put("date",df);
        fmts.put("list", df);
        fmts.put("map", df);
        byte[] file = SimpleExcelUtil.createExcel("测试", columnTitles, datas, columnFields,fmts);
        FileUtils.writeByteArrayToFile(new File("F:/excel.xls"), file);
        System.out.println("everything goes well");
    }
}

class MyJxlTestBean {
    private String str;
    private Date date;
    private int mint;
    private float mfloat;
    private List<Date> list;
    private Map<String,Date> map;

    public MyJxlTestBean(String str, Date date, int mint, float mfloat, List<Date> list, Map<String, Date> map) {
        super();
        this.str = str;
        this.date = date;
        this.mint = mint;
        this.mfloat = mfloat;
        this.list = list;
        this.map = map;
    }
    public String getStr() {
        return str;
    }
    public void setStr(String str) {
        this.str = str;
    }
    public Date getDate() {
        return date;
    }
    public void setDate(Date date) {
        this.date = date;
    }
    public int getMint() {
        return mint;
    }
    public void setMint(int mint) {
        this.mint = mint;
    }
    public float getMfloat() {
        return mfloat;
    }
    public void setMfloat(float mfloat) {
        this.mfloat = mfloat;
    }
    public List<Date> getList() {
        return list;
    }
    public void setList(List<Date> list) {
        this.list = list;
    }
    public Map<String, Date> getMap() {
        return map;
    }
    public void setMap(Map<String, Date> map) {
        this.map = map;
    }

}
