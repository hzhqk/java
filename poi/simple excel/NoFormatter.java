package excel;

/**
 * 不格式化
 *
 * @author hzhqk
 * @date 2020/06/08
 */
public class NoFormatter implements ExcelColumnFormatter {
    @Override
    public String format(Object o) {
        return null;
    }
}
