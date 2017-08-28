package com.fuzhu.base;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * Created by 符柱成 on 2017/8/24.
 */
public interface PoiInterface<T> {
    /*
     * 一些通用的方法：在此明确所有参数
     * int  excelVersion         excel的版本
     * String title           表格标题名
     * List<String> headersName      表格属性列名数组（即：每列标题）
     * List<String>  headersId        表格属性列名对应的字段（即：每列标题的英文标识--为了去list去拿）---你需要导出的字段名（所有接口都是支持headersId乱序的设计）
     * List<T> dtoList  或者    List<Map<String, Object>>  dtoList      想要导出的数据list（即：数据库查出的数据集合）   有两种风格：JavaBean风格  与  哈希数据结构风格
     * OutputStream out             与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     *
     * Workbook wb              Workbook工作簿对象
     * Sheet sheet           表格对象
     *
     * StyleInterface styleUtil       是我抽象出来的样式层，大家可继承ExcelStyleBase类实现自己的超高自定义样式
     *
     * int pageNum         分页码--针对大数据量的分页功能
     * int pageSize        每页的数量--针对大数据量的分页功能
     */


    /*
         （普通JavaBean结构）
     */
    //导出默认样式EXCEL文件--根据headersId筛选要导出的字段
    int exportBeanExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                           List<T> dtoList, OutputStream out);
    //导出自定义样式Excel文件--根据headersId筛选要导出的字段
    int exportStyleBeanExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                          List<T> dtoList, OutputStream out,StyleInterface styleUtil);
    //默认导出dtolist的所有数据--默认导出dtolist的所有数据
    int exportStyleBeanExcel(int excelVersion,String title, List<String> headersName,
                             List<T> dtoList, OutputStream out,StyleInterface styleUtil);

    //分页导出自定义样式Excel文件----拿到工作簿
    Workbook getPageExcelBook(int excelVersion);
    //分页导出自定义样式Excel文件----拿到表格
    Sheet getPageExcelSheet(Workbook wb,String bookTitle);
    //分页导出自定义样式Excel文件----标题栏
    Sheet exportPageTitleExcel(Workbook wb,Sheet sheet,List<String> headersName,StyleInterface styleUtil);


    //分页导出Bean结构自定义样式Excel文件----数据体
    Sheet exportPageContentBeanExcel(Workbook wb,Sheet sheet,List<String> headersId,List<T> dtoList,StyleInterface styleUtil,int pageNum,int pageSize);
    //分页导出Bean结构自定义样式Excel文件----数据体--没有标题栏字段匹配--默认导出dtolist的所有数据
    Sheet exportPageContentBeanExcel(Workbook wb,Sheet sheet,List<T> dtoList,StyleInterface styleUtil,int pageNum,int pageSize);


    //分页导出Map结构自定义样式Excel文件----数据体--根据headersId筛选要导出的字段
    Sheet exportPageContentMapExcel(Workbook wb,Sheet sheet,List<String> headersId,List<Map<String, Object>>  dtoList,StyleInterface styleUtil,int pageNum,int pageSize);
    //分页导出Map结构自定义样式Excel文件----数据体--没有标题栏字段匹配，数据体dtoList需要使用treemap。--默认导出dtolist的所有数据
    Sheet exportPageContentMapExcel(Workbook wb,Sheet sheet,List<Map<String, Object>>  dtoList,StyleInterface styleUtil,int pageNum,int pageSize);


    /*
        List<Map<String, Object>>结构
     */
    //导出默认样式的Map结构Excel--根据headersId筛选要导出的字段
    int exportMapExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                              List<Map<String, Object>> dtoList, OutputStream out) throws Exception ;
    //导出自定义样式的Map结构Excel--根据headersId筛选要导出的字段
    int exportStyleMapExcel(int excelVersion,String title, List<String> headersName, List<String> headersId,
                              List<Map<String, Object>> dtoList, OutputStream out,StyleInterface styleUtil) throws Exception ;
    //导出自定义样式的Map结构Excel--没有标题栏字段匹配，数据体dtoList需要使用treemap。--默认导出dtolist的所有数据
    int exportStyleMapExcel(int excelVersion,String title, List<String> headersName,
                                  List<Map<String, Object>> dtoList, OutputStream out,StyleInterface styleUtil) throws Exception ;
}
