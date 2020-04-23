package com.sinfeeloo.importassistant.entity;

import java.io.Serializable;

/**
 * Created by IntelliJ IDEA.
 *
 * @author : houly
 * Date: 14-7-15
 */

public class SheetInfo implements Serializable {
    private static final long serialVersionUID = 7843011810164624322L;
    private String sheetName;
    private int num = -1;

    public static long getSerialVersionUID() {
        return serialVersionUID;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getNum() {
        return num;
    }

    public void setNum(int num) {
        this.num = num;
    }
}
