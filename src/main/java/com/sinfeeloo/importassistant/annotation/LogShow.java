package com.sinfeeloo.importassistant.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 属性标注了该注解后比较实体时会比对改属性的值
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface LogShow {
    //比较结果字段描述
    public abstract String value() default "";
    //字段字典Map
    public abstract Dict[] dicts() default {};

    @Retention(RetentionPolicy.RUNTIME)
    @Target(ElementType.FIELD)
    public @interface Dict{
        public abstract String code() default "";
        public  abstract String text() default "";
    }



}
