package excel;

import com.jd.admin.car.util.DateUtils;

import java.util.Date;

/**
 * 默认日期格式化
 *
 * @author hzhqk
 * @date 2020/06/08
 */
public class DefaultDateFormatter implements ExcelColumnFormatter{
    @Override
    public String format(Object t) {
        Date date = (Date) t;
        return DateUtils.convertDateToString(date);
    }

}
