package com.fuzhu.util;

import com.fuzhu.base.PoiExcelBase;
import com.fuzhu.base.StyleInterface;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Created by 符柱成 on 2017/8/24.
 */
public class ExportMapExcel<T> extends PoiExcelBase<T> {


    /*
        导出默认样式EXCEL文件（根据headersId来导出对应字段，）--根据headersId筛选要导出的字段
    */
    @Override
    public int exportMapExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                            List<Map<String, Object>> dtoList, OutputStream out) throws Exception {
        int flag = 0;
        //写入excel
        Workbook wb = writeInExcel(excelVersion, title, headersName,headersId,dtoList,null);
        try {
            wb.write(out);
            flag = 1;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }

    /*
        导出自定义样式的Map结构Excel--根据headersId筛选要导出的字段
     */
    @Override
    public int exportStyleMapExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                                  List<Map<String, Object>> dtoList, OutputStream out,StyleInterface styleUtil) throws Exception {
        int flag = 0;
        //写入excel
        Workbook wb = writeInExcel(excelVersion, title, headersName,headersId,dtoList,styleUtil);
        try {
            wb.write(out);
            flag = 1;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }
    /*
           抽象出写入样式层
     */
    private Workbook writeInExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                                  List<Map<String, Object>> dtoList, StyleInterface styleUtil){

        //表头--标题栏
        Map<Integer, String> headersNameMap = getHeadersNameMap(headersName);
        //字段---标题的字段
        Map<Integer, String> titleFieldMap = getTitleFieldMap(headersId);
        // 声明一个工作薄03还是07版本
        Workbook wb = null;
        wb = getWorkbook(excelVersion);

        //获得表格并设置表格标题
        Sheet sheet =  wb.createSheet(title);
        //设置样式：
        CellStyle headerStyle = null;
        short rowHigh = 0;
        short columnWidth = 0;
        if (styleUtil!=null) {//策略是否为空
            headerStyle = styleUtil.setHeaderStyle(wb);//策略设置样式
            styleUtil.setRowHigh();//策略设置行高
            rowHigh = styleUtil.getRowHigh();//行高
            styleUtil.setColumnWidth();//策略设置列宽
            columnWidth = styleUtil.getColumnWidth();//列宽
        }
        if (headerStyle==null){
            headerStyle=getHeaderCellStyle(wb);
        }
        //行高--先设置行高再设置列宽
        if (rowHigh==0) {
            rowHigh = (short) (2 * 256);
        }
        sheet.setDefaultRowHeight(rowHigh);
        //列宽
        if (columnWidth==0) {
            columnWidth = 15;
        }
        sheet.setDefaultColumnWidth( columnWidth);
        CellStyle customizedStyle = null;
        if (styleUtil!=null) {
            //一个补偿方法，设定特定列宽
            styleUtil.setSpecifiedHighAndWidth(sheet);
            //如果有使用完全自定义方式，则覆盖上面所有的方式
            customizedStyle = styleUtil.setHeaderStyle(wb, sheet);
        }
        if (customizedStyle!=null) {
            headerStyle = customizedStyle;
        }
        //拿到第一行索引（标题栏）
        Row row = sheet.createRow(0);
        Cell cell = null;
        Collection c = headersNameMap.values();//拿到表格所有标题的value的集合
        Iterator<String> headersNameIt = c.iterator();//表格标题的迭代器
        //根据选择的字段生成表头--标题
        setTitle(row,headersNameIt,cell,headerStyle);
        //表格一行的字段的集合，以便拿到迭代器
        Collection zdC = titleFieldMap.values();
        Iterator<Map<String, Object>> titleFieldIt = dtoList.iterator();//总记录的迭代器
        //获取自定义的数据样式：
        CellStyle dataStyle = null;
        if (styleUtil!=null) {
            dataStyle = styleUtil.setDataStyle(wb);
        }
        if (dataStyle==null){
            dataStyle = getDataCellStyle(wb);
        }
        int zdRow = 0;//真正的数据记录的列序号
        doWriteInExcel(titleFieldIt,zdRow,sheet,zdC,dataStyle);
        return wb;
    }

    /*
        分页导出Map结构自定义样式Excel文件----数据体
     */
    @Override
    public Sheet exportPageContentMapExcel(Workbook wb, Sheet sheet, List<String> headersId, List<Map<String, Object>>  dtoList, StyleInterface styleUtil, int pageNum, int pageSize) {
        //字段---标题的字段
        Map<Integer, String> titleFieldMap = getTitleFieldMap(headersId);
        //表格一行的字段的集合
        Collection zdC = titleFieldMap.values();
        Iterator<Map<String, Object>> labIt = dtoList.iterator();//总记录的迭代器
        //获取自定义的数据样式：
        CellStyle dataStyle = null;
        dataStyle = styleUtil.setDataStyle(wb);
        if (dataStyle==null){
            dataStyle = getDataCellStyle(wb);
        }
        //写入excel
        writeInPageExcel(labIt, sheet, zdC,dataStyle,pageNum,pageSize);

        return sheet;
    }
    private void writeInPageExcel(Iterator<Map<String, Object>> titleFieldIt,Sheet sheet,Collection zdC,CellStyle dataStyle,int pageNum,int pageSize){
        int zdRow = (pageNum-1)*pageSize;//真正的数据记录的列序号
        doWriteInExcel(titleFieldIt,zdRow,sheet,zdC,dataStyle);
    }

    /*
        抽象出写入数据层---有标题字段匹对
    */
    private void doWriteInExcel(Iterator<Map<String, Object>> titleFieldIt,int zdRow,Sheet sheet,Collection zdC,CellStyle dataStyle){
        while (titleFieldIt.hasNext()) {//记录的迭代器，遍历总记录
            Map<String, Object> mapTemp = titleFieldIt.next();//拿到一条记录
            zdRow++;
            Row row = sheet.createRow(zdRow);
            int zdCell = 0;
            for (Map.Entry<String, Object> entry : mapTemp.entrySet()) {//支持headersId乱序的设计
                String key = entry.getKey();//记录的列字段
                Object value = entry.getValue();
                Iterator<String> zdIt = zdC.iterator();//一条记录的字段的集合的迭代器
                while (zdIt.hasNext()) {//遍历
                    String tempField = zdIt.next();//字段的暂存
                    if (key.equals(tempField) && value != null) {
                        Cell contentCell = row.createCell((short) zdCell);
                        contentCell.setCellValue(String.valueOf(value));//写进excel对象
                        contentCell.setCellStyle(dataStyle);
                        zdCell++;
                    }
                }
            }
        }
    }

    /*
        分页导出Map结构自定义样式Excel文件----数据体----没有标题栏字段匹配，数据体dtoList需要使用treemap。
     */
    @Override
    public Sheet exportPageContentMapExcel(Workbook wb, Sheet sheet, List<Map<String, Object>>  dtoList, StyleInterface styleUtil, int pageNum, int pageSize) {
        //表格一行的字段的集合
        Iterator<Map<String, Object>> labIt = dtoList.iterator();//总记录的迭代器
        //获取自定义的数据样式：
        CellStyle dataStyle = null;
        dataStyle = styleUtil.setDataStyle(wb);
        if (dataStyle==null){
            dataStyle = getDataCellStyle(wb);
        }
        //写入excel
        writeInPageExcel(labIt,sheet,dataStyle,pageNum,pageSize);
        return sheet;
    }

    //无标题字段匹对--分页
    private void writeInPageExcel(Iterator<Map<String, Object>> titleFieldIt,Sheet sheet ,CellStyle dataStyle,int pageNum,int pageSize){
        int zdRow = (pageNum-1)*pageSize;//真正的数据记录的列序号
        writeInExcelWithoutField(titleFieldIt,sheet,dataStyle,zdRow);
    }
    /*
        导出自定义样式的Map结构Excel--没有标题栏字段匹配，数据体dtoList需要使用treemap。--默认导出dtolist的所有数据
     */
    @Override
    public int exportStyleMapExcel(int excelVersion,String title, List<String> headersName,
                                         List<Map<String, Object>> dtoList, OutputStream out,StyleInterface styleUtil) throws Exception {
        int flag = 0;
        //写入excel
        Workbook wb = writeInExcelStyleWithoutField(excelVersion, title,headersName,dtoList,styleUtil);
        try {
            wb.write(out);
            flag = 1;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }
    //无标题字段匹对--不分页
    private Workbook writeInExcelStyleWithoutField(int excelVersion,String title, List<String> headersName,
                                               List<Map<String, Object>> dtoList, StyleInterface styleUtil){
        //表头--标题栏
        Map<Integer, String> headersNameMap = getHeadersNameMap(headersName);
        // 声明一个工作薄03还是07版本
        Workbook wb = null;
        wb = getWorkbook(excelVersion);

        //获得表格并设置表格标题
        Sheet sheet =  wb.createSheet(title);
        //设置样式：
        CellStyle headerStyle = null;
        headerStyle = styleUtil.setHeaderStyle(wb);
        if (headerStyle==null){
            headerStyle=getHeaderCellStyle(wb);
        }
        //行高--先设置行高再设置列宽
        styleUtil.setRowHigh();
        short rowHigh = styleUtil.getRowHigh();
        if (rowHigh==0) {
            rowHigh = (short) (2 * 256);
        }
        sheet.setDefaultRowHeight(rowHigh);
        //列宽
        styleUtil.setColumnWidth();
        short columnWidth = styleUtil.getColumnWidth();
        if (columnWidth==0) {
            columnWidth = 15;
        }
        sheet.setDefaultColumnWidth( columnWidth);
        //一个补偿方法，设定特定列宽
        styleUtil.setSpecifiedHighAndWidth(sheet);
        //如果有使用完全自定义方式，则覆盖上面所有的方式
        CellStyle customizedStyle = styleUtil.setHeaderStyle(wb,sheet);
        if (customizedStyle!=null) {
            headerStyle = customizedStyle;
        }
        //拿到第一行索引（标题栏）
        Row row = sheet.createRow(0);
        Cell cell = null;
        Collection c = headersNameMap.values();//拿到表格所有标题的value的集合
        Iterator<String> headersNameIt = c.iterator();//表格标题的迭代器
        //根据选择的字段生成表头--标题
        setTitle(row,headersNameIt,cell,headerStyle);
        //表格一行的字段的集合，以便拿到迭代器
        Iterator<Map<String, Object>> titleFieldIt = dtoList.iterator();//总记录的迭代器
        //获取自定义的数据样式：
        CellStyle dataStyle = null;
        dataStyle = styleUtil.setDataStyle(wb);
        if (dataStyle==null){
            dataStyle = getDataCellStyle(wb);
        }
        int zdRow = 1;//真正的数据记录的列序号
        writeInExcelWithoutField(titleFieldIt,sheet,dataStyle,zdRow);
        return wb;
    }
    /*
        抽象出写入数据层---无标题字段匹对，兼容分页与不分页
     */
    private void writeInExcelWithoutField(Iterator<Map<String, Object>> titleFieldIt, Sheet sheet,CellStyle dataStyle,int zdRow){
        while (titleFieldIt.hasNext()) {//记录的迭代器，遍历总记录
            Map<String, Object> mapTemp = titleFieldIt.next();//拿到一条记录
            zdRow++;
            Row row = sheet.createRow(zdRow);
            int zdCell = 0;
            if (mapTemp!=null) {
                Set<Map.Entry<String, Object>> entrySet = mapTemp.entrySet();
                for (Map.Entry<String, Object> entry : entrySet) {
                    //String key = entry.getKey();
                    Object value = entry.getValue();
                    Cell contentCell = row.createCell((short) zdCell);
                    contentCell.setCellValue(String.valueOf(value));//写进excel对象
                    contentCell.setCellStyle(dataStyle);
                    zdCell++;
                }
            }
        }
    }

}
