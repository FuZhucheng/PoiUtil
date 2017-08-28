package com.fuzhu.util;

/**
 * Created by 符柱成 on 2017/8/23.
 */

import com.fuzhu.base.PoiExcelBase;
import com.fuzhu.base.StyleInterface;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.*;

/**
 *            应用泛型，代表任意一个符合javabean风格的类
 *            注意这里为了简单起见，boolean型的属性xxx的get器方式为getXxx(),而不是isXxx()
 *            T这里代表一个不确定是实体类，即参数实体
 */
public  class ExportBeanExcel<T> extends PoiExcelBase<T> {

    /**
     * 这是一个通用的方法，利用了JAVA的反射机制，可以将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出
     *
     * title         表格标题名
     * headersName  表格属性列名数组
     * headersId    表格属性列名对应的字段
     *  dtoList     需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象
     *  out         与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     */
    @Override
    public  int exportBeanExcel(int excelVersion,String title, List<String> headersName,List<String> headersId,
                            List<T> dtoList, OutputStream out) {
        int flag = 0;
        //写入excel
        Workbook wb = writeInExcel( excelVersion, title,headersName, headersId,
                dtoList, out,null);
        try {
            wb.write(out);
            flag = 1;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return flag;
    }

    /*
        导出自定义样式Excel文件--根据headersId筛选要导出的字段
     */

    @Override
    public int exportStyleBeanExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                                 List<T> dtoList, OutputStream out,StyleInterface styleUtil) {
        int flag = 0;
        //写入excel
        Workbook wb = writeInExcel( excelVersion, title,headersName, headersId,
                dtoList, out,styleUtil);
        try {
            wb.write(out);
            flag = 1;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }
    /*
          导出自定义样式Excel文件--默认导出dtoList所有字段
     */
    @Override
    public int exportStyleBeanExcel(int excelVersion,String title, List<String> headersName,
                                    List<T> dtoList, OutputStream out,StyleInterface styleUtil) {
        int flag = 0;
        //写入excel
        Workbook wb = writeInExcel( excelVersion, title,headersName, null,
                dtoList, out,styleUtil);
        try {
            wb.write(out);
            flag = 1;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return flag;
    }
    /*
       抽象出写入样式层---有标题栏字段匹配与兼容无标题栏字段匹配的情况
    */
    private Workbook writeInExcel(int excelVersion,String title, List<String> headersName,List<String> headersId,
                                  List<T> dtoList, OutputStream out,StyleInterface styleUtil){
        //表头--标题栏
        Map<Integer, String> headersNameMap = getHeadersNameMap(headersName);
        if (headersId==null){//兼容无标题栏字段匹配的情况
            headersId = new ArrayList();
            Field[] fields = dtoList.get(0).getClass().getDeclaredFields();
            int i = 0;
            while(i<fields.length) {
                Field field = fields[i];
                String fieldName = field.getName();//属性名
                headersId.add(fieldName);
                i++;
            }
        }
        //字段---标题的字段
        Map<Integer, String> titleFieldMap = getTitleFieldMap(headersId);
        // 声明一个工作薄
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
            headerStyle=getHeaderCellStyle(wb);//拿默认样式
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
        sheet.setDefaultColumnWidth(columnWidth);
        //一个补偿方法，设定特定列宽
        CellStyle customizedStyle = null;
        if (styleUtil!=null) {
            styleUtil.setSpecifiedHighAndWidth(sheet);
            //如果有使用完全自定义方式，则覆盖上面所有的方式
            customizedStyle = styleUtil.setHeaderStyle(wb, sheet);
        }
        if (customizedStyle!=null) {
            headerStyle = customizedStyle;
        }

        Row row = sheet.createRow(0);
        Cell cell = null;
        Collection c = headersNameMap.values();//拿到表格所有标题的value的集合
        Iterator<String> headersNameIt = c.iterator();//表格标题的迭代器
        //根据选择的字段生成表头--标题
        setTitle(row,headersNameIt,cell,headerStyle);
        /* ---------------------------以上是标题栏，以下是数据列-----------------------------   */
        //表格一行的字段的集合
        Collection zdC = titleFieldMap.values();
        Iterator<T> labIt = dtoList.iterator();//总记录的迭代器
        //获取自定义的数据样式：
        CellStyle dataStyle = null;
        if (styleUtil!=null) {
            dataStyle = styleUtil.setDataStyle(wb);
        }
        if (dataStyle==null){
            dataStyle = getDataCellStyle(wb);
        }
        int zdRow =0;//列序号
        writeInExcel(labIt,sheet,zdC,dataStyle,zdRow);
        return wb;
    }


    /*
        分页导出Bean结构自定义样式Excel文件----数据体
     */
    @Override
    public  Sheet exportPageContentBeanExcel(Workbook wb,Sheet sheet,List<String> headersId,List<T> dtoList,StyleInterface styleUtil,int pageNum,int pageSize){

        //写入excel
        writeInPageExcel(wb,sheet,headersId,dtoList,styleUtil,pageNum,pageSize);
        return sheet;
    }
    /*
           分页导出Bean结构自定义样式Excel文件----数据体--没有标题栏字段匹配--默认导出dtolist的所有数据
     */
    @Override
    public  Sheet exportPageContentBeanExcel(Workbook wb,Sheet sheet,List<T> dtoList,StyleInterface styleUtil,int pageNum,int pageSize){
        List<String> headersId = null;
        if (headersId==null){//兼容无标题栏字段匹配的情况
            headersId = new ArrayList();
            Field[] fields = dtoList.get(0).getClass().getDeclaredFields();
            int i = 0;
            while(i<fields.length) {
                Field field = fields[i];
                String fieldName = field.getName();//属性名
                headersId.add(fieldName);
                i++;
            }
        }
        //写入excel
        writeInPageExcel(wb,sheet,headersId,dtoList,styleUtil,pageNum,pageSize);

        return sheet;
    }
    /*
        分页导出Bean结构自定义样式Excel文件---样式层
     */
    private void writeInPageExcel(Workbook wb,Sheet sheet,List<String> headersId,List<T> dtoList,StyleInterface styleUtil,int pageNum,int pageSize){

        //字段---标题的字段
        Map<Integer, String> titleFieldMap = getTitleFieldMap(headersId);
        //表格一行的字段的集合
        Collection zdC = titleFieldMap.values();
        Iterator<T> labIt = dtoList.iterator();//总记录的迭代器
        //获取自定义的数据样式：
        CellStyle dataStyle = null;
        dataStyle = styleUtil.setDataStyle(wb);
        if (dataStyle==null){
            dataStyle = getDataCellStyle(wb);
        }

        int zdRow =(pageNum-1)*pageSize;//列序号
        writeInExcel(labIt,sheet,zdC,dataStyle,zdRow);
    }

    //抽象出写入数据层
    private void writeInExcel(Iterator<T> labIt,Sheet sheet,Collection zdC,CellStyle dataStyle,int zdRow){
        while (labIt.hasNext()) {//记录的迭代器，遍历总记录
            int zdCell = 0;
            zdRow++;
            Row row = sheet.createRow(zdRow);
            T l = (T) labIt.next();
            // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
            Field[] fields = l.getClass().getDeclaredFields();//获得JavaBean全部属性
            for (short i = 0; i < fields.length; i++) {//遍历属性，比对
                Field field = fields[i];
                String fieldName = field.getName();//属性名
                Iterator<String> zdIt = zdC.iterator();//一条字段的集合的迭代器
                while (zdIt.hasNext()) {//遍历要导出的字段集合
                    if (zdIt.next().equals(fieldName)) {//比对JavaBean的属性名，一致就写入，不一致就丢弃
                        String getMethodName = "get"
                                + fieldName.substring(0, 1).toUpperCase()
                                + fieldName.substring(1);//拿到属性的get方法
                        Class tCls = l.getClass();//拿到JavaBean对象
                        try {
                            Method getMethod = tCls.getMethod(getMethodName,
                                    new Class[] {});//通过JavaBean对象拿到该属性的get方法，从而进行操控
                            Object val = getMethod.invoke(l, new Object[] {});//操控该对象属性的get方法，从而拿到属性值
                            String textVal = null;
                            if (val!= null) {
                                textVal = String.valueOf(val);//转化成String
                            }else{
                                textVal = "";
                            }
                            Cell contentCell = row.createCell((short) zdCell);
                            contentCell.setCellValue(textVal);//写进excel对象
                            contentCell.setCellStyle(dataStyle);
                            zdCell++;
                        } catch (SecurityException e) {
                            e.printStackTrace();
                        } catch (IllegalArgumentException e) {
                            e.printStackTrace();
                        } catch (NoSuchMethodException e) {
                            e.printStackTrace();
                        } catch (IllegalAccessException e) {
                            e.printStackTrace();
                        } catch (InvocationTargetException e) {
                            e.printStackTrace();
                        }catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                }
            }
        }
    }
}

