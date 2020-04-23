package com.sinfeeloo.importassistant.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by pc on 14-6-12.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelTitle {
    public abstract String titleName() default "";
    public abstract int index() default -1;
    public abstract int width() default -1;
    public abstract String type() default "String";
    /**
     * 格式，默认是""暂且只处理date类型的,其他类型不处理
     * date类型未指定格式时格式为yyyy/MM/dd
     * @return
     */
    public abstract String format() default "";
    public abstract int  rowSpan() default 1;//合并行(跨越几行，默认1行)
    public abstract int groupNumber()  default 1;//分组，默认第一组
}
