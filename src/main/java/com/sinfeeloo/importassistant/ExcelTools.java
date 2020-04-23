package com.sinfeeloo.importassistant;

import com.google.common.base.Strings;
import com.google.common.io.Files;
import com.intelliquor.cloud.shop.common.annotation.LevelExcelTitle;
import com.intelliquor.cloud.shop.common.entity.*;
import com.intelliquor.cloud.shop.common.template.AbstractExcelExportTemplate;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;


/**
 * Excel操作工具类
 * Created by lanjsh
 * 2016年4月19日15:18:15
 */
public final class ExcelTools extends AbstractExcelExportTemplate {
    private ExcelTools() {
    }

    private ExcelInfo excelInfo;

    public void setExcelInfo(ExcelInfo excelInfo) {
        this.excelInfo = excelInfo;
    }

    @Override
    protected void buildBody(int sheetIndex) {

        Sheet sheet = getSheet(sheetIndex);

        int startIndex = this.getBodyStartIndex(sheetIndex);

        //查看是否需要自己画title
        List<LevelTitle> levelTitles = excelInfo.getLevelTitles();


        if (levelTitles != null && levelTitles.size() != 0) {
            //需要自己画title
            startIndex = drawTitle(sheetIndex, levelTitles, new String[][]{excelInfo.getTitleArray()}, new Integer[][]{excelInfo.getRowSpansArray()});
            //startIndex = drawTitle(sheetIndex,levelTitles,excelInfo.getTitleInfo());

        }

        //设置列宽
        for (int i = 0; i < excelInfo.getColWidths().size(); i++) {
            int width = excelInfo.getColWidths().get(i);
            if (width != -1) {
                sheet.setColumnWidth(i, excelInfo.getColWidths().get(i));
            }

        }


        int curRow = startIndex;
        int curCel = 0;
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        short df = sheet.getWorkbook().createDataFormat().getFormat("yyyy/MM/dd");
        cellStyle.setDataFormat(df);
        List<List<ExcelBodyInfo>> datas = excelInfo.getBodyInfo();

        //判断sheet显示条数
        List<SheetInfo> sheetInfos = excelInfo.getSheetInfos();
        if (sheetInfos == null || sheetInfos.size() == 0) {
            for (List<ExcelBodyInfo> data : datas) {
                Row row = sheet.createRow(curRow);
                for (ExcelBodyInfo excelBodyInfo : data) {
                    row.createCell(curCel++).setCellValue(excelBodyInfo.getValue());
                }
                curRow += 1;
                curCel = 0;
            }
        } else {
            //获得显示条数范围
            int begin = 0;
            int end = 0;
            for (int i = 0; i < sheetInfos.size(); i++) {
                SheetInfo sheetInfo = sheetInfos.get(i);
                if (i == 0) {
                    end += sheetInfo.getNum() - 1;
                } else {
                    begin = end + 1;
                    end = begin + sheetInfo.getNum() - 1;
                }


                if (i == sheetIndex) {
                    break;
                }
            }


            if ((datas.size() - 1) >= (begin)) {
                if (end > (datas.size() - 1)) {
                    end = datas.size() - 1;
                }
                List<List<ExcelBodyInfo>> datasSub = datas.subList(begin, end + 1);
                for (List<ExcelBodyInfo> data : datasSub) {
                    Row row = sheet.createRow(curRow);
                    for (ExcelBodyInfo excelBodyInfo : data) {
                        row.createCell(curCel);
                        if ("Integer".equals(excelBodyInfo.getType())) {
                            row.getCell(curCel).setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            row.getCell(curCel).setCellValue(Integer.valueOf(excelBodyInfo.getValue()));
                        } else if ("Double".equals(excelBodyInfo.getType())) {
                            row.getCell(curCel).setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            row.getCell(curCel).setCellValue(Double.valueOf(excelBodyInfo.getValue()));
                        } else if ("BigDecimal".equals(excelBodyInfo.getType())) {
                            row.getCell(curCel).setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                            row.getCell(curCel).setCellValue(new BigDecimal(excelBodyInfo.getValue()).doubleValue());
                        } else if ("Date".equals(excelBodyInfo.getType())) {
                            try {
                                String format = "yyyy/MM/dd";
                                if (StringUtils.isNoneBlank(excelBodyInfo.getFormat())) {
                                    format = excelBodyInfo.getFormat();
                                    CellStyle customerCellStyle = sheet.getWorkbook().createCellStyle();
                                    short customerDf = sheet.getWorkbook().createDataFormat().getFormat(format);
                                    customerCellStyle.setDataFormat(customerDf);
                                    row.getCell(curCel).setCellStyle(customerCellStyle);
                                } else {
                                    row.getCell(curCel).setCellStyle(cellStyle);
                                }
                                row.getCell(curCel).setCellValue(new SimpleDateFormat(format).parse(excelBodyInfo.getValue()));
                            } catch (Exception e) {
                                row.getCell(curCel).setCellValue(excelBodyInfo.getValue());
                            }

                        } else if ("String".equals(excelBodyInfo.getType())) {
                            if (StringUtils.isNoneBlank(excelBodyInfo.getFormat())) {
                                String format = excelBodyInfo.getFormat();
                                CellStyle customerCellStyle = sheet.getWorkbook().createCellStyle();
                                short customerDf = sheet.getWorkbook().createDataFormat().getFormat(format);
                                customerCellStyle.setDataFormat(customerDf);
                                row.getCell(curCel).setCellStyle(customerCellStyle);
                            }
                            row.getCell(curCel).setCellValue(excelBodyInfo.getValue());
                        } else {
                            row.getCell(curCel).setCellValue(excelBodyInfo.getValue());
                        }
                        curCel++;
                    }
                    curRow += 1;
                    curCel = 0;
                }

            }


        }


    }

    @Override
    public String[] getSheetNames() {
        List<SheetInfo> sheetInfos = excelInfo.getSheetInfos();
        if (sheetInfos != null && sheetInfos.size() != 0) {
            String[] sheetNames = new String[sheetInfos.size()];
            for (int i = 0; i < sheetInfos.size(); i++) {
                SheetInfo sheetInfo = sheetInfos.get(i);
                sheetNames[i] = sheetInfo.getSheetName();
            }
            return sheetNames;
        } else {
            return new String[]{Files.getNameWithoutExtension(excelInfo.getFileName())};
        }

    }

    @Override
    public String[][] getTitles() {
        List<LevelTitle> levelTitles = excelInfo.getLevelTitles();
        if (levelTitles != null && levelTitles.size() != 0) {
            return new String[][]{};
        } else {
            return new String[][]{excelInfo.getTitleArray()};
        }

    }


    private int drawTitle(int sheetIndex, List<LevelTitle> levelTitles, String[][] titles, Integer[][] rowSpans) {
        Sheet sheet = this.getSheet(sheetIndex);
        int titleStartIndex = this.getTitleStartIndex(sheetIndex);

//        for(int i = 0; i < ts.length; i++){
//            sheet.setColumnWidth(i, columnWidth);
//            createStyledCell(rowTitle, i, ts[i], this.titleRowStyle);
//        }

        for (LevelTitle t : levelTitles) {
            Row rowTitle = sheet.createRow(titleStartIndex);
            t.getColSpans();
            int colindex = 0;
            LevelExcelTitle.Align[] aligns = new LevelExcelTitle.Align[t.getTitleNames().length];
            if (t.getAligns() == null || t.getAligns().length == 0) {
                for (int i = 0; i < aligns.length; i++) {
                    aligns[i] = LevelExcelTitle.Align.CENTER;
                }
            } else if (t.getAligns().length < aligns.length) {
                for (int i = 0; i < aligns.length; i++) {
                    if (i <= (aligns.length - 1)) {
                        aligns[i] = t.getAligns()[i];
                    } else {
                        aligns[i] = LevelExcelTitle.Align.CENTER;
                    }
                }
            } else {
                aligns = t.getAligns();
            }


            for (int i = 0; i < t.getTitleNames().length; i++) {
                String titleName = t.getTitleNames()[i];

                Cell cell = rowTitle.createCell(colindex);
                cell.setCellValue(titleName);
                cell.setCellStyle(crateMergeTitleCellStyle());
                //默认合并1列
                int col = 1;
                if (t.getColSpans() != null && t.getColSpans().length > 0) {
                    if (t.getColSpans()[i] > 1) {
                        col = t.getColSpans()[i];
                    }
                }
                int endColindex = colindex + col - 1;
                //默认合并1行
                int rows = 1;
                if (t.getRowSpans() != null && t.getRowSpans().length > 0) {
                    if (t.getRowSpans()[i] > 1) {
                        rows = t.getRowSpans()[i];
                    }
                }
                //当前行+要合并的行数-1(减1是因为行号是从0开始的)
                int endRow = titleStartIndex + rows - 1;
                //四个参数:起始行，结束行，起始列，结束列
                CellRangeAddress region = new CellRangeAddress(titleStartIndex, endRow, colindex, endColindex);
                sheet.addMergedRegion(region);

                int border = 1;
                RegionUtil.setBorderBottom(border, region, sheet, this.workbook);
                RegionUtil.setBorderLeft(border, region, sheet, this.workbook);
                RegionUtil.setBorderRight(border, region, sheet, this.workbook);
                RegionUtil.setBorderTop(border, region, sheet, this.workbook);

                colindex = endColindex + 1;
            }
            titleStartIndex++;

        }
        this.titles = titles;
        this.rowSpans = rowSpans;
//        buildTitle(sheetIndex);

        //Row rowTitle = sheet.createRow(titleStartIndex);
        //合并行后先判断有没有创建行，如已创建，则使用，如没创建则创建
        Row rowTitle = sheet.getRow(titleStartIndex);
        if (rowTitle == null) {
            rowTitle = sheet.createRow(titleStartIndex);
        }
        String[] ts = this.titles[0];
        Integer[] rspan = this.rowSpans[0];
        for (int i = 0; i < ts.length; i++) {
            //四个参数:起始行，结束行，起始列，结束列
            int rowspan = rspan[i];
            if (rowspan == 1) {
                sheet.setColumnWidth(i, columnWidth);
                //createStyledCell(rowTitle, i, ts[i], this.titleRowStyle);
                createStyledCell(rowTitle, i, ts[i], crateMergeTitleCellStyle());
            }
        }
        titleStartIndex++;
        //titleStartIndex = titleStartIndex+ titles.length;
        return titleStartIndex;
    }

    /**
     * 导出文件(自定义文件名)
     *
     * @param list
     * @param FileInName 导出文件名
     * @return
     * @throws IOException
     */
    public static FileItem exportByFile(List list, String FileInName, int groupId) throws IOException {
        ExcelInfo excelInfo = AnnotationUtil.getExcelInfo(list, groupId);
        FileItem fileItem = new FileItem();
        if (StringUtils.isNotBlank(FileInName)) {
            fileItem.setFileName(FileInName);
        } else {
            fileItem.setFileName(excelInfo.getFileName());
        }
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ExcelTools.export(excelInfo, baos);
        fileItem.setContent(baos.toByteArray());
        return fileItem;
    }

    /**
     * 根据数据和分组导出数据
     *
     * @param list
     * @param groupNumber 分组编号
     * @return
     * @throws IOException
     */
    public static FileItem exportByFile(List list, int groupNumber) throws IOException {
        ExcelInfo excelInfo = AnnotationUtil.getExcelInfo(list, groupNumber);
        /*FileItem fileItem=new FileItem();
        fileItem.setFileName(excelInfo.getFileName());
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        com.sinfeeloo.importUtils.ExcelTools.export(excelInfo, baos);
        fileItem.setContent(baos.toByteArray());
        return fileItem;*/
        return ExcelTools.exportByExcelInfo(excelInfo);
    }

    /**
     * 导出文件
     *
     * @param list
     * @return
     * @throws IOException
     */
    public static FileItem exportByFile(List list) throws IOException {
        ExcelInfo excelInfo = AnnotationUtil.getExcelInfo(list, 1);
        /*FileItem fileItem=new FileItem();
        fileItem.setFileName(excelInfo.getFileName());
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        com.sinfeeloo.importUtils.ExcelTools.export(excelInfo, baos);
        fileItem.setContent(baos.toByteArray());
        return fileItem;*/
        return ExcelTools.exportByExcelInfo(excelInfo);
    }

    public static FileItem exportByExcelInfo(ExcelInfo excelInfo) throws IOException {
        FileItem fileItem = new FileItem();
        fileItem.setFileName(excelInfo.getFileName());
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ExcelTools.export(excelInfo, baos);
        fileItem.setContent(baos.toByteArray());
        return fileItem;
    }

    public static void export(ExcelInfo excelInfo, OutputStream outputStream) throws IOException {
        ExcelTools excel = new ExcelTools();
        excel.setExcelInfo(excelInfo);
        excel.doExport(outputStream, null);
    }


    public static Response<ExcelInfo> transferStreamToExcel(InputStream inputStream, int sheetIndex, String dateFormat) {
        Response<ExcelInfo> result = new Response<ExcelInfo>();
        ExcelInfo excelinfo = new ExcelInfo();
        List<ExcelTitleInfo> titleInfo = new ArrayList<ExcelTitleInfo>();
        List<List<ExcelBodyInfo>> body = new ArrayList<List<ExcelBodyInfo>>();
        XSSFWorkbook workbook = null;
        XSSFSheet sheet;
        XSSFRow row;

        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (Exception e) {

            return result;
        }
        //读取表头 begin
        sheet = workbook.getSheetAt(sheetIndex);
        row = sheet.getRow(0);
        // 标题总列数
        int colNum = row.getPhysicalNumberOfCells();
        //System.out.println("colNum:" + colNum);
        for (int i = 0; i < colNum; i++) {
            Cell cell = row.getCell(i);
            if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                result.setError("导入出错，请不要修改单元格格式(" + i + ")");
                return result;
            }
            ExcelTitleInfo titleE = new ExcelTitleInfo();
            titleE.setTitle(cell.getStringCellValue());
            titleInfo.add(titleE);
        }
        //读取表头 end
        //读取内容 begin
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = 1; i <= rowNum; i++) {
            List<ExcelBodyInfo> rowResult = new ArrayList<ExcelBodyInfo>();
            row = sheet.getRow(i);
            int j = 0;
            while (j < colNum) {
                Cell cell = row.getCell(j);
                String cellValue = getCellValue(i, j, cell, result, dateFormat);
                if (cellValue == null) {
                    return result;
                }
                ExcelBodyInfo excelBodyInfo = new ExcelBodyInfo();
                excelBodyInfo.setValue(cellValue);
                excelBodyInfo.setType("String");
                rowResult.add(excelBodyInfo);
                j++;
            }
            body.add(rowResult);
        }
        //读取内容 end

        excelinfo.setTitleInfo(titleInfo);
        excelinfo.setBodyInfo(body);
        result.setResult(excelinfo);

        return result;
    }

    private static String getCellValue(int line, int col, Cell cell, Response response, String dateFormat) {
        String result = null;
        if (cell == null) {
            result = "";
        } else {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        if (Strings.isNullOrEmpty(dateFormat)) {
                            dateFormat = "yyyy-MM-dd";
                        }
                        result = TimeUtilis.getDateStringByForm(cell.getDateCellValue(), dateFormat);
                    } else {
                        BigDecimal big = new BigDecimal(cell.getNumericCellValue());
                        result = big.toString();
                        //解决1234.0  去掉后面的.0
                        if (null != result && !"".equals(result.trim())) {
                            String[] item = result.split("[.]");
                            if (1 < item.length && "0".equals(item[1])) {
                                result = item[0];
                            }
                        }
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = String.valueOf(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    //result=cell.getCellFormula();
                    try {
                        result = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        try {
                            result = String.valueOf(cell.getNumericCellValue());
                        } catch (Exception ex) {
                            result = "";
                        }
                    }
                    break;
                case Cell.CELL_TYPE_BLANK:
                    result = "";
                    break;
                case Cell.CELL_TYPE_ERROR:
                    result = "";
                    break;
                default:
                    response.setError("获取内容发生错误！行号：" + line + " 列号:" + col);
                    break;
            }
        }

        return result;
    }

    public static Response<List<ExcelInfo>> transferStreamToExcelByAllSheets(InputStream inputStream) {
        Response<List<ExcelInfo>> result = new Response<List<ExcelInfo>>();
        List<ExcelInfo> excelinfoList = new ArrayList<ExcelInfo>();
        XSSFWorkbook workbook = null;
        XSSFSheet sheet;
        XSSFRow row;

        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (Exception e) {

            return result;
        }
        for (int k = 0; k < workbook.getNumberOfSheets(); k++) {
            ExcelInfo excelinfo = new ExcelInfo();
            List<List<ExcelBodyInfo>> body = new ArrayList<List<ExcelBodyInfo>>();
            List<ExcelTitleInfo> titleInfo = new ArrayList<ExcelTitleInfo>();

            //读取表头 begin
            sheet = workbook.getSheetAt(k);
            String sheetName = sheet.getSheetName();
            row = sheet.getRow(0);
            // 标题总列数
            int colNum = row.getPhysicalNumberOfCells();
            //System.out.println("colNum:" + colNum);
            for (int i = 0; i < colNum; i++) {
                Cell cell = row.getCell(i);
                if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                    result.setError("导入出错，请不要修改单元格格式(" + i + ")");
                    return result;
                }
                ExcelTitleInfo titleE = new ExcelTitleInfo();
                titleE.setTitle(cell.getStringCellValue());
                titleInfo.add(titleE);
            }
            //读取表头 end
            //读取内容 begin
            // 得到总行数
            int rowNum = sheet.getLastRowNum();
            // 正文内容应该从第二行开始,第一行为表头的标题
            for (int i = 1; i <= rowNum; i++) {
                List<ExcelBodyInfo> rowResult = new ArrayList<ExcelBodyInfo>();
                row = sheet.getRow(i);
                int j = 0;
                while (j < colNum) {
                    if (null == row) {
                        j++;
                        continue;
                    }
                    Cell cell = row.getCell(j);
                    String cellValue = getCellValue(i, j, cell, result);
                    if (cellValue == null) {
                        return result;
                    }
                    ExcelBodyInfo excelBodyInfo = new ExcelBodyInfo();
                    excelBodyInfo.setValue(cellValue);
                    excelBodyInfo.setType("String");
                    rowResult.add(excelBodyInfo);
                    j++;
                }
                body.add(rowResult);
            }
            //读取内容 end

            excelinfo.setTitleInfo(titleInfo);
            excelinfo.setBodyInfo(body);
            excelinfo.setFileName(sheetName);
            excelinfoList.add(excelinfo);
        }
        result.setResult(excelinfoList);

        return result;
    }

    public static Response<ExcelInfo> transferStreamToExcel(InputStream inputStream) {
        Response<ExcelInfo> result = new Response<ExcelInfo>();
        ExcelInfo excelinfo = new ExcelInfo();
        List<ExcelTitleInfo> titleInfo = new ArrayList<ExcelTitleInfo>();
        List<List<ExcelBodyInfo>> body = new ArrayList<List<ExcelBodyInfo>>();
        XSSFWorkbook workbook = null;
        XSSFSheet sheet;
        XSSFRow row;

        try {
            workbook = new XSSFWorkbook(inputStream);
        } catch (Exception e) {

            return result;
        }
        //读取表头 begin
        sheet = workbook.getSheetAt(0);
        row = sheet.getRow(0);
        // 标题总列数
        int colNum = row.getPhysicalNumberOfCells();
        //System.out.println("colNum:" + colNum);
        for (int i = 0; i < colNum; i++) {
            Cell cell = row.getCell(i);
            if (cell.getCellType() != Cell.CELL_TYPE_STRING) {
                result.setError("导入出错，请不要修改单元格格式(" + i + ")");
                return result;
            }
            ExcelTitleInfo titleE = new ExcelTitleInfo();
            titleE.setTitle(cell.getStringCellValue());
            titleInfo.add(titleE);
        }
        //读取表头 end
        //读取内容 begin
        // 得到总行数
        int rowNum = sheet.getLastRowNum();
        int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
        // 正文内容应该从第二行开始,第一行为表头的标题
        for (int i = 1; i <= rowNum; i++) {
            List<ExcelBodyInfo> rowResult = new ArrayList<ExcelBodyInfo>();
            row = sheet.getRow(i);
            if(null!=row){
                int j = 0;
                while (j < colNum) {
                    Cell cell = row.getCell(j);
                    String cellValue = getCellValue(i, j, cell, result);
                    if (cellValue == null) {
                        return result;
                    }
                    ExcelBodyInfo excelBodyInfo = new ExcelBodyInfo();
                    excelBodyInfo.setValue(cellValue);
                    excelBodyInfo.setType("String");
                    rowResult.add(excelBodyInfo);
                    j++;
                }
                body.add(rowResult);
            }
        }
        //读取内容 end

        excelinfo.setTitleInfo(titleInfo);
        excelinfo.setBodyInfo(body);
        result.setResult(excelinfo);

        return result;
    }


    private static String getCellValue(int line, int col, Cell cell, Response response) {
        String result = null;
        if (cell == null) {
            result = "";
        } else {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                    result = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {
                        result = TimeUtilis.getDateStringByForm(cell.getDateCellValue(), "yyyy-MM-dd");
                    } else {
                        BigDecimal big = new BigDecimal(cell.getNumericCellValue());
                        result = big.toString();
                        //解决1234.0  去掉后面的.0
                        if (null != result && !"".equals(result.trim())) {
                            String[] item = result.split("[.]");
                            if (1 < item.length && "0".equals(item[1])) {
                                result = item[0];
                            }
                        }
                    }
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    result = String.valueOf(cell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    //result=cell.getCellFormula();
                    try {
                        result = String.valueOf(cell.getStringCellValue());
                    } catch (IllegalStateException e) {
                        try {
                            result = String.valueOf(cell.getNumericCellValue());
                        } catch (Exception ex) {
                            result = "";
                        }
                    }
                    break;
                case Cell.CELL_TYPE_BLANK:
                    result = "";
                    break;
                case Cell.CELL_TYPE_ERROR:
                    result = "";
                    break;
                default:
                    response.setError("获取内容发生错误！行号：" + line + " 列号:" + col);
                    break;
            }
        }

        return result;
    }

    /**
     * 创建单元格,文本,数字,日期
     *
     * @param row
     * @param index
     * @param cellValue
     * @param cellStyle
     * @param type
     */
    protected void createStyledCell(Row row, int index, String cellValue, CellStyle cellStyle, String type) {
        if (Strings.isNullOrEmpty(type)) {
            type = "String";
        }
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Cell cell = row.createCell(index);
        if ("BigDecimal".equals(type)) {
            try {
                cell.setCellValue(Double.parseDouble(cellValue));
            } catch (Exception e) {
                cell.setCellValue(cellValue);
            }

        } else if ("Integer".equals(type)) {
            try {
                cell.setCellValue(Integer.parseInt(cellValue));
            } catch (Exception e) {
                cell.setCellValue(cellValue);
            }

        } else if ("Date".equals(type)) {
            try {
                cell.setCellValue(sdf.parse(cellValue));
            } catch (ParseException e) {
                cell.setCellValue(cellValue);
            }
        } else {
            cell.setCellValue(cellValue);
        }

        cell.setCellStyle(cellStyle);
    }
}
