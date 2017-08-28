package com.fuzhu.test;

import com.fuzhu.base.PoiExcelBase;
import com.fuzhu.base.PoiInterface;
import com.fuzhu.base.StyleInterface;
import com.fuzhu.model.Student;
import com.fuzhu.styleImpl.MyStyle;
import com.fuzhu.util.PoiBeanFactory;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by 符柱成 on 2017/8/25.
 */
public class ExportPageBeanTest {
    public static void main(String [] args) throws Exception {

        List<String> listName = new ArrayList<>();
        listName.add("id");
        listName.add("名字");
        listName.add("性别");
        List<String> listId = new ArrayList<>();
        listId.add("id");
        listId.add("name");
        listId.add("sex");


        FileOutputStream exportXls = null;
        if (PoiExcelBase.EXCEL_VERSION_07==0) {
            exportXls = new FileOutputStream("E://工单信息表PageNoHeaders.xls");
        }else {
            exportXls = new FileOutputStream("E://工单信息表PageNoHeaders.xlsx");
        }

        /*
            （一）去工厂拿导出工具
         */
        PoiInterface<Student> poiInterface = PoiBeanFactory.getInstance().getPoiUtil(PoiExcelBase.EXPORT_SIMPLE_EXCEL);
        /*
            （二）拿到工作簿对象（可选版本）
         */
        Workbook workbook = poiInterface.getPageExcelBook(PoiExcelBase.EXCEL_VERSION_07);
        /*
            （三）拿到表格对象（填写表格名字）
         */
        Sheet sheet = poiInterface.getPageExcelSheet(workbook,"测试工作簿的title");
        /*
            （四）自定义样式
         */
        StyleInterface myStyle = new MyStyle();
        /*
            （五）导出标题栏数据先
         */
        sheet = poiInterface.exportPageTitleExcel(workbook,sheet,listName,myStyle);
        /*
            （六）分页导出数据列（注意控制好页码以及一页的数量--做过分页功能的应该都有这个经验的）
         */
        int q=0;
        for (int t =1;t<6;t++){
            List<Student> list = new ArrayList<>();
            list.add(new Student(++q,"张三asdf","男"+t));
            list.add(new Student(++q,"李四asd","男"+t));
            list.add(new Student(++q,"王五bhasdcfvbhujidsaub","女"+t));
            //默认导出全部数据
            poiInterface.exportPageContentBeanExcel(workbook,sheet,list,myStyle,t,3);
            //根据listId导出数据
            //poiInterface.exportPageContentBeanExcel(workbook,sheet,listId,list,myStyle,t,3);
        }
        /*
              （七）写入到流对象
         */
        workbook.write(exportXls);
        exportXls.close();

    }
}
