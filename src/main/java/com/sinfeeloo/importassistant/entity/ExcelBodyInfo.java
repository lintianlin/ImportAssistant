package com.sinfeeloo.importassistant.entity;

import org.springframework.beans.factory.annotation.Value;

import java.io.Serializable;

/**
 * Created by pc on 14-6-12.
 */

public class ExcelBodyInfo extends Object implements Serializable
{
    private String value;
    @Value("String")
    private String type;
    /**
     * 格式
     */
    private String format;

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }
}
