package com.sinfeeloo.importassistant.entity;

import com.google.common.collect.Lists;

import java.util.List;

/**
 * Created by pc on 14-6-12.
 */

public class ExcelInfo {

    private String fileName;

    private List<SheetInfo> sheetInfos;

    private List<LevelTitle> levelTitles;

    private List<ExcelTitleInfo> titleInfo;

    private List<List<ExcelBodyInfo>> bodyInfo;



    private List<String> titles;
    private String[] titlesArray;
    private List<Integer> widths;
    private List<Integer> rowSpans;
    private Integer[] rowSpansArray;

    private int groupId;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public List<SheetInfo> getSheetInfos() {
        return sheetInfos;
    }

    public void setSheetInfos(List<SheetInfo> sheetInfos) {
        this.sheetInfos = sheetInfos;
    }

    public List<LevelTitle> getLevelTitles() {
        return levelTitles;
    }

    public void setLevelTitles(List<LevelTitle> levelTitles) {
        this.levelTitles = levelTitles;
    }

    public List<ExcelTitleInfo> getTitleInfo() {
        return titleInfo;
    }

    public void setTitleInfo(List<ExcelTitleInfo> titleInfo) {
        this.titleInfo = titleInfo;
    }

    public List<List<ExcelBodyInfo>> getBodyInfo() {
        return bodyInfo;
    }

    public void setBodyInfo(List<List<ExcelBodyInfo>> bodyInfo) {
        this.bodyInfo = bodyInfo;
    }

    public List<String> getTitles() {
        return titles;
    }

    public void setTitles(List<String> titles) {
        this.titles = titles;
    }

    public String[] getTitlesArray() {
        return titlesArray;
    }

    public void setTitlesArray(String[] titlesArray) {
        this.titlesArray = titlesArray;
    }

    public List<Integer> getWidths() {
        return widths;
    }

    public void setWidths(List<Integer> widths) {
        this.widths = widths;
    }

    public List<Integer> getRowSpans() {
        return rowSpans;
    }

    public void setRowSpans(List<Integer> rowSpans) {
        this.rowSpans = rowSpans;
    }

    public int getGroupId() {
        return groupId;
    }

    public void setGroupId(int groupId) {
        this.groupId = groupId;
    }

    public List<String> getTitle(){
        if(this.titleInfo == null){
            return null;
        }
        if(titles!=null){
            return titles;
        }
        List<String> list = Lists.newArrayList();
        for(ExcelTitleInfo e :titleInfo){
           list.add(e.getTitle());
        }
        titles = list;
        return list;
    }

    public String[] getTitleArray(){
        List<String> list = getTitle();
        if(list == null){
            return null;
        }
        if(titlesArray!=null){
            return titlesArray;
        }
        String[] array = new String[list.size()];
        titlesArray = list.toArray(array);
        return titlesArray;
    }


    public List<Integer> getColWidths(){
        if(this.titleInfo == null){
            return null;
        }
        if(widths!=null){
            return widths;
        }
        List<Integer> list = Lists.newArrayList();
        for(ExcelTitleInfo e :titleInfo){
            list.add(e.getWidth());
        }
        widths = list;
        return list;
    }

    public List<Integer> getRowSpansList(){
        if(this.titleInfo == null){
            return null;
        }
        if(rowSpans!=null){
            return rowSpans;
        }
        List<Integer> list = Lists.newArrayList();
        for(ExcelTitleInfo e :titleInfo){
            list.add(e.getRowSpan());
        }
        rowSpans = list;
        return list;
    }

    public Integer[] getRowSpansArray(){
        if(this.titleInfo == null){
            return null;
        }
        if(rowSpansArray!=null){
            return rowSpansArray;
        }
        List<Integer> list = getRowSpansList();
        Integer[] array = new Integer[list.size()];
        rowSpansArray = list.toArray(array);
        return rowSpansArray;
    }

}
