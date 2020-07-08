package excel;

import java.lang.annotation.*;

/**
 * excel列名
 *
 * @author hzhqk
 * @date 2020/06/08
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {
    /**
     * 列名/数据项
     * @return
     */
    String name() default "";

    /**
     * 列顺序，值越小越排在前面
     * @return
     */
    int order() default 0;

    /**
     * 列格式化器
     * @return
     */
    Class<? extends ExcelColumnFormatter> formatter() default NoFormatter.class;
}
