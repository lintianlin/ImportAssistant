package com.sinfeeloo.importassistant.entity;







import com.intelliquor.cloud.shop.common.annotation.LevelExcelTitle;

import java.io.Serializable;

/**
 * Created by IntelliJ IDEA.
 * User: houly
 * Date: 14-7-15
 */
public class LevelTitle extends Object implements Serializable,Comparable<LevelTitle> {
    private static final long serialVersionUID = -4165837812436350098L;

    private String[] titleNames;
    private int titleLevel;
    private int[] rowSpans;//合并行
    private int[] colSpans;//合并列
    private LevelExcelTitle.Align[] aligns;

    private int  groupId; //分组编号

    public static long getSerialVersionUID() {
        return serialVersionUID;
    }

    public String[] getTitleNames() {
        return titleNames;
    }

    public void setTitleNames(String[] titleNames) {
        this.titleNames = titleNames;
    }

    public int getTitleLevel() {
        return titleLevel;
    }

    public void setTitleLevel(int titleLevel) {
        this.titleLevel = titleLevel;
    }


    public LevelExcelTitle.Align[] getAligns() {
        return aligns;
    }

    public void setAligns(LevelExcelTitle.Align[] aligns) {
        this.aligns = aligns;
    }

    public int[] getRowSpans() {
        return rowSpans;
    }

    public void setRowSpans(int[] rowSpans) {
        this.rowSpans = rowSpans;
    }

    public int[] getColSpans() {
        return colSpans;
    }

    public void setColSpans(int[] colSpans) {
        this.colSpans = colSpans;
    }

    public int getGroupId() {
        return groupId;
    }

    public void setGroupId(int groupId) {
        this.groupId = groupId;
    }

    @Override
    public int compareTo(LevelTitle o) {
        if(this.titleLevel == o.getTitleLevel()){
            return 0;
        }

        else if(this.titleLevel>  o.getTitleLevel()){
            return 1;
        }
        else {
            return -1;
        }
    }

    public boolean equals(Object o){
        if(o instanceof LevelTitle){
            return this.titleLevel == ((LevelTitle)o).titleLevel;
        }else{
            return false;
        }
    }

    public int hashCode(){
        return this.titleLevel;
    }
}
