package com.fuzhu.styleImpl;

import com.fuzhu.base.ExcelStyleBase;
import com.fuzhu.base.StyleInterface;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by 符柱成 on 2017/8/25.
 */
public class MyStyle extends ExcelStyleBase {

    @Override
    public CellStyle setHeaderStyle(Workbook wb) {

        // 设置字体
        Font font = wb.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 10);
        //字体加粗
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");
        font.setColor((short) 10);
        // 生成一个样式
        CellStyle headerStyle = wb.createCellStyle();
        // 设置背景色
        headerStyle.setFillForegroundColor((short) 13);
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //设置字体
        headerStyle.setFont(font);
        //设置顶边框;
        headerStyle.setBorderTop(CellStyle.BORDER_THIN);
        //设置右边框;
        headerStyle.setBorderRight(CellStyle.BORDER_THIN);
        //设置左边框;
        headerStyle.setBorderLeft(CellStyle.BORDER_THIN);
        //设置底边框;
        headerStyle.setBorderBottom(CellStyle.BORDER_THIN);
        //设置自动换行;
        headerStyle.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐;
        headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        return headerStyle;
    }

    @Override
    public CellStyle setDataStyle(Workbook wb) {
        // 设置字体
        Font font = wb.createFont();
        //设置字体大小
        font.setFontHeightInPoints((short) 10);
        //字体加粗
//        font.setBoldweight(Font.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");
        font.setColor((short) 32767);
        // 生成一个样式
        CellStyle dataStyle = wb.createCellStyle();
        // 设置背景色
        dataStyle.setFillForegroundColor((short)70);
        dataStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        //设置字体
        dataStyle.setFont(font);
        //设置顶边框;
        dataStyle.setBorderTop(CellStyle.BORDER_THIN);
        //设置右边框;
        dataStyle.setBorderRight(CellStyle.BORDER_THIN);
        //设置左边框;
        dataStyle.setBorderLeft(CellStyle.BORDER_THIN);
        //设置底边框;
        dataStyle.setBorderBottom(CellStyle.BORDER_THIN);
        //设置自动换行;
        dataStyle.setWrapText(false);
        //设置水平对齐的样式为居中对齐;
        dataStyle.setAlignment(CellStyle.ALIGN_CENTER);
        //设置垂直对齐的样式为居中对齐;
        dataStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        return dataStyle;
    }

    @Override
    public void setRowHigh() {
        setMyRowHigh((short) (2 * 256));
    }

    @Override
    public void setColumnWidth() {
        setMyColumnWidth((short) 30);
    }

    @Override
    public Map<Integer,Integer> setMySpecifiedHighAndWidth() {
        Map<Integer,Integer> map  = new HashMap<>();
        map.put(1,3000);
        map.put(2,3000);
        return map;
    }
    @Override
    public CellStyle setHeaderStyle(Workbook wb, Sheet sheet) {
        return null;
    }
}
