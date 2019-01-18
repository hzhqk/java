import com.alibaba.otter.canal.protocol.CanalEntry;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.tqmall.lsc.common.tools.DateUtils;
import com.tqmall.lsc.mq_canallog.CanalLogMsgProcessor;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang.StringUtils;
import org.springframework.core.convert.converter.Converter;
import org.springframework.core.convert.support.DefaultConversionService;

import javax.annotation.PostConstruct;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

public abstract class AbstractCanalLogMsgProcessor implements CanalLogMsgProcessor {

    private DefaultConversionService conversionService = new DefaultConversionService() {
        {
            addConverter(new Converter<String, Date>() {
                @Override
                public Date convert(String source) {
                    if (StringUtils.isBlank(source)) {
                        return null;
                    }
                    return DateUtils.convertStringToDate(source);
                }
            });
        }
    };

    private ConcurrentHashMap<String, Map<String, Field>> cachedClzFields = new ConcurrentHashMap<>();


    /**
     * 获取改变<b>前后</b>的数据，将canal消息数据转为bean（只赋值private、public、protected属性，不赋值static、final等其他属性）
     * 注意：属性名不能包含特殊字符
     *
     * @param rowChange
     * @param clz
     * @return
     * @throws IllegalAccessException 参数为空时会抛异常
     * @throws InstantiationException
     */
    public <T> List<RowDataPair<T>> getChanges(CanalEntry.RowChange rowChange, Class<T> clz) throws InstantiationException, IllegalAccessException {
        if (rowChange == null || clz == null || rowChange.getRowDatasList() == null) {
            throw new IllegalArgumentException("rowChange or clz can't be empty.");
        }
        Map<String, Field> beanFields = getClzFields(clz);
        List<RowDataPair<T>> result = Lists.newArrayList();
        for (CanalEntry.RowData rowData : rowChange.getRowDatasList()) {
            T dataBefore = convertRowData(rowData.getBeforeColumnsList(), beanFields, clz);
            T dataAfter = convertRowData(rowData.getAfterColumnsList(), beanFields, clz);
            result.add(new RowDataPair<>(dataBefore, dataAfter));
        }
        return result;
    }

    /**
     * 获取改变<b>前</b>的数据，将canal消息数据转为bean（只赋值private、public、protected属性，不赋值static、final等其他属性）
     * 注意：属性名不能包含特殊字符
     *
     * @param rowChange
     * @param clz
     * @return
     * @throws IllegalAccessException 参数为空时会抛异常
     * @throws InstantiationException
     */
    public <T> List<T> getChangesBefore(CanalEntry.RowChange rowChange, Class<T> clz) throws IllegalAccessException, InstantiationException {
        return getChangesBeforeOrAfter(rowChange, clz, true);
    }

    /**
     * 获取改变<b>后</b>的数据，将canal消息数据转为bean（只赋值private、public、protected属性，不赋值static、final等其他属性）
     * 注意：属性名不能包含特殊字符
     *
     * @param rowChange
     * @param clz
     * @return
     * @throws IllegalAccessException 参数为空时会抛异常
     * @throws InstantiationException
     */
    public <T> List<T> getChangesAfter(CanalEntry.RowChange rowChange, Class<T> clz) throws IllegalAccessException, InstantiationException {
        return getChangesBeforeOrAfter(rowChange, clz, false);
    }

    /**
     * bean初始化完成后注册需要的转换器，已默认添加Date（格式：yyyy-MM-dd HH:mm:ss）转换器
     * 在方法体内，如下使用：
     * <pre>
     *    addConverter(new Converter<String, Date>() {
     *
     *    });
     * </pre>
     */
    @PostConstruct
    protected void registerConverters() {

    }

    /**
     * 给class对应field设置别名
     */
    @PostConstruct
    protected void aliasClzFields() {

    }

    /**
     * 给field名称设置别名，将忽略大小写并去除下划线
     * 不建议业务逻辑中设置别名，请重写aliasClzFields方法
     * <pre>
     * 如：canal msg：  {“birth_day”:"1970-01-01 00:00:00"}
     *    bean中属性为 birth
     *    那么 aliasField(Bean.class, birth, birth_day) 或者 aliasField(Bean.class, birth, birthday)等
     * </pre>
     *
     * @param clz
     * @param originName
     * @param aliasName
     * @param <T>
     */
    protected <T> void aliasField(Class<T> clz, String originName, String aliasName) {
        Map<String, Field> clzFields = getClzFields(clz);
        aliasName = aliasName.toLowerCase().replace("_", "");
        clzFields.put(aliasName, clzFields.get(originName.toLowerCase()));
    }

    /**
     * 注册自定义配置
     * 不可直接调用，请重写registerConverters()
     *
     * @param converter
     * @see com.tqmall.lsc.mq_canallog.impl.AbstractCanalLogMsgProcessor#registerConverters()
     */
    protected void addConverter(Converter converter) {
        conversionService.addConverter(converter);
    }

    private <T> List<T> getChangesBeforeOrAfter(CanalEntry.RowChange rowChange, Class<T> clz, boolean isBefore) throws InstantiationException, IllegalAccessException {
        if (rowChange == null || clz == null || rowChange.getRowDatasList() == null) {
            throw new IllegalArgumentException("rowChange or clz can't be empty.");
        }
        Map<String, Field> beanFields = getClzFields(clz);
        List<T> result = Lists.newArrayList();
        for (CanalEntry.RowData rowData : rowChange.getRowDatasList()) {
            List<CanalEntry.Column> columnsList = isBefore ? rowData.getBeforeColumnsList() : rowData.getAfterColumnsList();
            T data = convertRowData(columnsList, beanFields, clz);
            result.add(data);
        }
        return result;
    }

    private <T> Map<String, Field> getClzFields(Class<T> clz) {
        Map<String, Field> beanFields = cachedClzFields.get(clz.getName());
        if (beanFields == null || beanFields.size() <= 0) {
            beanFields = getAllFieldsForBean(clz);
            cachedClzFields.putIfAbsent(clz.getName(), beanFields);
            beanFields = cachedClzFields.get(clz.getName());
        }
        return beanFields;
    }

    private <T> T convertRowData(List<CanalEntry.Column> cols, Map<String, Field> beanFields, Class<T> clz) throws IllegalAccessException, InstantiationException {
        if (CollectionUtils.isEmpty(cols)) {
            return null;
        }
        T bean = clz.newInstance();
        for (CanalEntry.Column col : cols) {
            String name = col.getName().toLowerCase().replace("_", "");
            String value = col.getValue();
            Field field = beanFields.get(name);
            if (field == null) {
                continue;
            }
            field.set(bean, value == null ? null : conversionService.convert(value, field.getType()));
        }
        return bean;
    }

    private <T> Map<String, Field> getAllFieldsForBean(Class<T> clz) {
        Map<String, Field> result = Maps.newHashMap();
        Class tmpClz = clz;
        // 不获取Object层的属性
        String finalParent = "java.lang.object";
        while (tmpClz != null && !tmpClz.getName().toLowerCase().equals(finalParent)) {
            // 只获取bean普通属性
            for (Field field : tmpClz.getDeclaredFields()) {
                // 不在设置数据时设置访问权限
                field.setAccessible(true);
                int modifiers = field.getModifiers();
                if (modifiers == Modifier.PUBLIC || modifiers == Modifier.PRIVATE || modifiers == Modifier.PROTECTED) {
                    result.put(field.getName().toLowerCase(), field);
                }
            }
            tmpClz = tmpClz.getSuperclass();
        }
        return result;
    }

    @Getter
    @Setter
    public static class RowDataPair<T> {
        private T before;
        private T after;

        public RowDataPair(T before, T after) {
            this.before = before;
            this.after = after;
        }
    }
}
