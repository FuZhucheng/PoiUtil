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
import java.util.*;

/**
 * Created by 符柱成 on 2017/8/26.
 */
public class ExportPageMapTest {
    public static void main(String [] args) throws Exception {
        Map<String, String> mapT = new HashMap<String, String>();
        mapT.put("1", "value1");
        mapT.put("2", "value2");
        mapT.put("3", "value3");



        List<String> listName = new ArrayList<>();
        listName.add("id");
        listName.add("名字");
        listName.add("性别");
        List<String> listId = new ArrayList<>();
        listId.add("sex");
        listId.add("name");
        listId.add("id");
//        List<Map<String,Object>> listBa = new ArrayList<>();
//        for (int t=0;t<6;t++){
//            Map<String,Object> map =  new TreeMap();
//            map.put("id",""+t);
//            map.put("name","abc"+t);
//            map.put("sex","男"+t);
//            listBa.add(map);
//        }
        FileOutputStream exportXls = null;
        if (PoiExcelBase.EXCEL_VERSION_07==0) {
            exportXls = new FileOutputStream("E://工单信息表PageMap--有标题字典匹对.xls");
        }else {
            exportXls = new FileOutputStream("E://工单信息表PageMap--有标题字典匹对.xlsx");
        }

        /*
            （一）去工厂拿导出工具
         */
        PoiInterface<Student> poiInterface = PoiBeanFactory.getInstance().getPoiUtil(PoiExcelBase.EXPORT_MAP_EXCEL);
        /*
            （二）拿到工作簿对象（可选版本）
         */
        Workbook workbook = poiInterface.getPageExcelBook(PoiExcelBase.EXCEL_VERSION_07);
        /*
            （三）拿到表格对象（填写表格名字）
         */
        Sheet sheet = poiInterface.getPageExcelSheet(workbook,"测试工作簿的title");
        /*
            （四）自定义样式（可无）
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
            List<Map<String,Object>> listB = new ArrayList<>();
            for (int p=0;p<6;p++){
                Map<String,Object> map = new TreeMap<>();
                q++;
                map.put("id",q);
                map.put("name","abc"+t);
                map.put("sex","男"+t);
                listB.add(map);
            }
           // poiInterface.exportPageContentMapExcel(workbook,sheet,listB,myStyle,t,6);
            poiInterface.exportPageContentMapExcel(workbook,sheet,listId,listB,myStyle,t,6);
        }
        /*
              （七）写入到流对象
         */
        workbook.write(exportXls);
        exportXls.close();
    }
}
