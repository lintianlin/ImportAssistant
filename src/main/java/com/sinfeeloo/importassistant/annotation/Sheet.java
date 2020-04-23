package com.sinfeeloo.importassistant.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by IntelliJ IDEA.
 * User: houly
 * Date: 14-7-15
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Sheet {

    //默认的sheet名称是sheet+num 即 sheet1
    public abstract String[] sheetNames();
    public abstract int numPerSheet() default  -1;
    public abstract int[] numOfSheet() default {};
    public abstract int   groupNumber() default 1;//分组编号，如果一个实体类要导出多张报表时用，默认1
}
