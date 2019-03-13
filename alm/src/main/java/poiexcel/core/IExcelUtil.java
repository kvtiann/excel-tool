package poiexcel.core;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * excel 导入导出工具
 *
 * @param <T>
 * @author 神盾局
 * @date 2016年8月9日 下午5:30:23
 */
public interface IExcelUtil<T> {

    /**
     * 构建一个导入导出工具
     *
     * @param clazz 一个类类型
     * @return IExcelUtil<T> 返回类型
     */
    IExcelUtil<T> build(Class<T> clazz);

    /**
     * 数据导出
     *
     * @param sheetName
     * @param input
     * @return List<T> 导出数据
     */
    List<T> importExcel(String sheetName, InputStream input);


    boolean exportExcel(List<T> list, String sheetName, List<String> titles, OutputStream output);


    boolean exportExcel(List<T> lists[], String sheetNames[], Map<Integer, List<String>> titleMap, OutputStream output);

}
