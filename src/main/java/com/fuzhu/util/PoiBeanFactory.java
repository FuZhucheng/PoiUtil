package com.fuzhu.util;


import com.fuzhu.base.PoiExcelBase;
import com.fuzhu.base.PoiInterface;

/**
 * Created by 符柱成 on 2017/8/24.
 */
public class PoiBeanFactory<T extends PoiExcelBase> {


    //静态内部类懒汉单例
    private PoiBeanFactory(){
    }
    public static synchronized final PoiBeanFactory getInstance() {
        return SingletonHolder.INSTANCE;
    }
    //静态内部类
    private static class SingletonHolder {
        static final PoiBeanFactory INSTANCE = new PoiBeanFactory();
    }
    //根据状态获得想要的实例
    public PoiInterface<T> getPoiUtil(int type){
        PoiInterface<T> poiInterface = null;
        switch (type){
            case -1:
                poiInterface = new ExportBeanExcel();
                break;
            case -2:

                break;
            case -3:
                break;
            case -4:
                break;
            case -5:
                poiInterface = new ExportMapExcel();
                break;
            case -6:
                break;
        }
        return  poiInterface;
    }

}
