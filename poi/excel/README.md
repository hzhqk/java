#### V2.0 说明

此版本基于注解更方便使用，使用示例见main方法。createExcel方法不再返回byte数组，如需返回数组可自己封装或调用后拿到HssfWorkbook后使用以下操作来确保相关的流可以正常关闭
```
获取
excel bytes示例：```
```
     ByteArrayOutputStream os = new ByteArrayOutputStream();
     workbook.write(os);
    或直接使用 ExcelUtil.getExcelBytes(hssfWorkbook);
```
返回HssfWorkbook目的是为了在生成的Excel基础上再次创建新的sheet页并写入数据

依赖的jar至少有：
```
<!-- poi excel 导出 -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>4.0.1</version>
        </dependency>
        <dependency>
            <groupId>com.google.guava</groupId>
            <artifactId>guava</artifactId>
            <version>19.0</version>
        </dependency>
        <dependency>
            <groupId>org.apache.commons</groupId>
            <artifactId>commons-lang3</artifactId>
            <version>3.6</version>
        </dependency>
        <!-- 如果不想依赖这个，那就去除write2Response方法 -->
        <dependency>
            <groupId>javax.servlet</groupId>
            <artifactId>servlet-api</artifactId>
            <version>2.4</version>
            <scope>provided</scope>
        </dependency>
```
