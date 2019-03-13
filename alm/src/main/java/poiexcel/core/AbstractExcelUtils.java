package poiexcel.core;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * 常用 抽象方法实现
 *
 * 
 * @param <T>
 */
public abstract class AbstractExcelUtils<T> extends AbstractExcelUtil<T> {



	@Override
	public List<T> importExcel(String sheetName, InputStream input) {
		List<T> list = null;
        try {
			HSSFWorkbook workbook = new HSSFWorkbook(input);
			HSSFSheet sheet = workbook.getSheet(sheetName);
			if (!sheetName.trim().equals("")) {
				// 如果指定sheet名,则取指定sheet中的内容.
			    sheet = workbook.getSheet(sheetName);
			}  
			if (sheet == null) {
				// 如果传入的sheet名不存在则默认指向第1个sheet.
			    sheet = workbook.getSheetAt(0);
			}		
			list = dispatch(sheet);
		} catch (IOException e) {
			e.printStackTrace();
		}

        return list;  
	}

	@SuppressWarnings("unchecked")
	@Override
	public boolean exportExcel(List<T> list, String sheetName, List<String> titles , OutputStream output) {
		//此处 对类型进行转换
        List<T> ilist = new ArrayList<>();
        for (T t : list) {
            ilist.add(t);
        }
        List<T>[] lists = new ArrayList[1];  
        lists[0] = ilist;  

        String[] sheetNames = new String[1];  
        sheetNames[0] = sheetName;
		Map<Integer,List<String>> map = new HashMap<>(1);
		map.put(0,titles);
		return exportExcel(lists, sheetNames,map, output);
	}

	@Override
	public boolean exportExcel(List<T>[] lists, String[] sheetNames, Map<Integer,List<String>> titleMap, OutputStream output) {
		if (lists.length != sheetNames.length) {
			System.out.println("数组长度不一致");
			return false;
		}

		// 创建excel工作簿
		ExcelTool tool = new ExcelTool();
		HSSFWorkbook wb = tool.createWorkbook();
		// 创建第一个sheet（页），命名为 new sheet
		for (int ii = 0; ii < lists.length; ii++) {
			List<T> list = lists[ii];
			// 产生工作表对象			
			HSSFSheet sheet = tool.createSheet(ii, sheetNames[ii]);

			// 创建表头
			createHeader(tool ,wb, sheet,titleMap.get(ii));
			// 写入数据
			int rowStart = 1;
			if (null != titleMap.get(ii)) {
				rowStart += titleMap.get(ii).size();
			}
			for (T t : list) {
				createRow(tool,t, sheet,wb);
				rowStart = mergedRegio(t, sheet,wb, rowStart);
			}
			
		}
		try {
			output.flush();
			wb.write(output);
			output.close();
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			System.out.println("Output is closed ");
			return false;
		}

	}
	
}
