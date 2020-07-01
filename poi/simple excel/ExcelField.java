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
    String name() default "";
    Class<? extends ExcelColumnFormatter> formatter() default NoFormatter.class;
}
