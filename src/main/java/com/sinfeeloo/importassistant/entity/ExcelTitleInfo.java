package com.sinfeeloo.importassistant.entity;

import java.io.Serializable;

/**
 * Created by pc on 14-6-12.
 *
 * @author CHQ
 */
public class ExcelTitleInfo extends Object implements Comparable<ExcelTitleInfo>, Serializable {
    private int index;
    private String title;
    private int width;
    private String type;
    /**
     * 格式
     */
    private String format;
    private String fieldName;
    /**
     * 合并行，跨越几行
     */
    private int rowSpan;
    /**
     * 分组编号
     */
    private int groupNumber;

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public int getWidth() {
        return width;
    }

    public void setWidth(int width) {
        this.width = width;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getFieldName() {
        return fieldName;
    }

    public void setFieldName(String fieldName) {
        this.fieldName = fieldName;
    }

    public int getRowSpan() {
        return rowSpan;
    }

    public void setRowSpan(int rowSpan) {
        this.rowSpan = rowSpan;
    }

    public int getGroupNumber() {
        return groupNumber;
    }

    public void setGroupNumber(int groupNumber) {
        this.groupNumber = groupNumber;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    @Override
    public int compareTo(ExcelTitleInfo o) {
        if (this.index == o.getIndex()) {
            return 0;
        } else if (this.index > o.getIndex()) {
            return 1;
        } else {
            return -1;
        }
    }

    @Override
    public boolean equals(Object o) {
        if (o instanceof ExcelTitleInfo) {
            return this.index == ((ExcelTitleInfo) o).getIndex();
        } else {
            return false;
        }
    }

    @Override
    public int hashCode() {
        return this.index;
    }
}
