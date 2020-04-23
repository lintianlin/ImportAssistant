package com.sinfeeloo.importassistant.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by pc on 14-6-12.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {
    public abstract String fileName();
    public abstract String fileExtension() default "xlsx";
    public abstract int    groupNumber() default 1;//分组编号，如果一个实体类要导出多张报表时用，默认1
}
