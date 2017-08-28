package com.fuzhu.base;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by 符柱成 on 2017/8/24.
 */
public interface StyleInterface {
    //设置标题栏的样式
    CellStyle setHeaderStyle(Workbook wb);
    //设置数据列的样式
    CellStyle setDataStyle(Workbook wb);
    //设置行高（自动设置每一行）
    void setRowHigh();
    //设置列宽（自动设置每一列）
    void setColumnWidth();
    //可利用此方法设定特定的列宽与行高---这个方法请不要覆写或重载，这个是给抽象类以及底层封装使用的
    void setSpecifiedHighAndWidth(Sheet sheet);

    /*
        当你使用以下这个方法的sheet对象时，请不要使用上面的setHeaderStyle(Workbook wb)、setRowHigh()、setColumnWidth()、setSpecifiedHighAndWidth(Sheet sheet)方法。因为下面是完全自定义，会完全覆盖上面方法的。
        同时请小心使用sheet对象，此处调用及其容易破坏封装。
     */
    //高度自定义标题栏样式--可以针对单列单行 宽高
    CellStyle setHeaderStyle(Workbook wb, Sheet sheet);


    //获取行高
    short getRowHigh();
    //获取列宽
    short getColumnWidth();
}
