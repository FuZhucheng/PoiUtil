package com.fuzhu.test;

import com.fuzhu.base.PoiExcelBase;
import com.fuzhu.base.PoiInterface;
import com.fuzhu.base.StyleInterface;
import com.fuzhu.model.Student;
import com.fuzhu.styleImpl.MyStyle;
import com.fuzhu.util.PoiBeanFactory;

import java.io.FileOutputStream;
import java.util.*;

/**
 * Created by 符柱成 on 2017/8/25.
 */
public class ExportMapTest {
    public static void main(String [] args) throws Exception {

        List<String> listName = new ArrayList();
        listName.add("id");
        listName.add("名字");
        listName.add("性别");
        List<String> listId = new ArrayList();
        listId.add("id");
        listId.add("sex");
        listId.add("name");
        List<Map<String,Object>> listB = new ArrayList();
        for (int t=0;t<6;t++){
            Map<String,Object> map = new TreeMap();
            map.put("id",t);
            map.put("name","abc"+t);
            map.put("sex","男"+t);
            listB.add(map);
        }
        FileOutputStream exportXls = null;
        if (PoiExcelBase.EXCEL_VERSION_07==0) {
            exportXls = new FileOutputStream("E://工单信息表Map.xls");
        }else {
            exportXls = new FileOutputStream("E://工单信息表Map.xlsx");
        }
        /*
            （一）去工厂拿导出工具
         */
        PoiInterface<Student> poiInterface = PoiBeanFactory.getInstance().getPoiUtil(PoiExcelBase.EXPORT_MAP_EXCEL);
        /*
            （二）自定义样式（可无）
         */
        StyleInterface myStyle = new MyStyle();
        /*
            （三）根据需求选择接口方法（返回码：1是成功，0为失败）
         */
        //导出默认样式的Map结构Excel--根据headersId筛选要导出的字段
        //int flag = poiInterface.exportMapExcel(PoiExcelBase.EXCEL_VERSION_07,"测试POI导出EXCEL文档",listName,listId,listB,exportXls);

        //导出自定义样式的Map结构Excel--根据headersId筛选要导出的字段
        //int flag = poiInterface.exportStyleMapExcel(PoiExcelBase.EXCEL_VERSION_07,"测试POI导出EXCEL文档",listName,listId,listB,exportXls,myStyle);
        //导出自定义样式的Map结构Excel--没有标题栏字段匹配，数据体dtoList需要使用treemap。--默认导出dtolist的所有数据
        int flag = poiInterface.exportStyleMapExcel(PoiExcelBase.EXCEL_VERSION_07,"测试POI导出EXCEL文档",listName,listB,exportXls,myStyle);

        System.out.println("flag  : "+flag);
        exportXls.close();
    }
}
