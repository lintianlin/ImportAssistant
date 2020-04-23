package com.sinfeeloo.importassistant.annotation;

/**
 * Created by IntelliJ IDEA.
 * User: houly
 * Date: 14-7-15
 */

public @interface LevelExcelTitle {
    public abstract String[] titleNames();
    public abstract int titleLevel();
    public abstract int[] rowSpans() default {}; //合并行，如果不填，默认合并一行
    public abstract int[] colSpans() default {}; //如果不填写每个都是默认占一列
    public abstract Align[] aligns() default {}; //默认居中
    public abstract int  groupNumber()  default 1;  //分组编号，默认第一组


    public static enum Align{
        LEFT,RIGHT,CENTER
    }
}
