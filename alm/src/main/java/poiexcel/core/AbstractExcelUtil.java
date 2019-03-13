package poiexcel.core;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.util.List;

/**
 * 
* 定义需要重写的细节
* 
* @param <T>
 */
public abstract class AbstractExcelUtil<T> implements IExcelUtil<T>{
	
	public abstract void createHeader(ExcelTool tool, HSSFWorkbook wb, HSSFSheet sheet, List<String> titles);
	
	public abstract void createRow(ExcelTool tool, Object t, HSSFSheet sheet, HSSFWorkbook wb);
	
	public abstract int mergedRegio(Object t, HSSFSheet sheet, HSSFWorkbook wb, int rowStart);
	
	/**
	 * 
	* 负责调度 将excel 数据转化为list
	* @param sheet
	* @return  List<?> 返回类型  
	* @throws
	 */
	public abstract List<T> dispatch(HSSFSheet sheet);

}
