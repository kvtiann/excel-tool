package poiexcel.core;


import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import poiexcel.annotation.ExcelAttribute;
import poiexcel.annotation.ExcelElement;
import poiexcel.annotation.ExcelID;
import poiexcel.config.ElementTypePath;


import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 简单的实现类
 *
 * @param <T>
 * @author 神盾局
 * @date 2016年8月10日 上午10:33:18
 */
public class ExcelUtils<T> extends AbstractExcelUtilss<T> {
    private Logger logger = Logger.getLogger(ExcelUtils.class);

    private Class<T> clazz;

    private HSSFRow row;

    private boolean flag;

    {
        flag = true;
    }

    @Override
    public IExcelUtil<T> build(Class<T> clazz) {
        this.clazz = clazz;
        return this;
    }

    @Override
    public void createHeader(ExcelTool tool, HSSFWorkbook wb, HSSFSheet sheet, List<String> titles) {

        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
        style.setFillBackgroundColor(HSSFColor.GREY_40_PERCENT.index);
        style.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框

        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.GREY_80_PERCENT.index);//HSSFColor.VIOLET.index //字体颜色
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//粗体显示
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 12);//设置字体大小
        style.setFont(font);

        // 得到所有的字段
        List<Field> fields = getAllField(clazz, null);
        if (null != titles && titles.size() > 0) {
            for (int i = 0; i < titles.size(); i++) {
                HSSFRow titleRow = tool.createRow(sheet);
                titleRow.setHeight((short) (2 * 256));
                HSSFCell titleCell = titleRow.createCell(0);// 创建列
                titleCell.setCellValue(titles.get(i));
                titleCell.setCellStyle(style);
                CellRangeAddress cra = new CellRangeAddress(i, i, 0, fields.size() - 1);
                sheet.addMergedRegion(cra);
                // 使用RegionUtil类为合并后的单元格添加边框
                RegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN, cra, sheet, wb); // 下边框
                RegionUtil.setBorderLeft(HSSFCellStyle.BORDER_THIN, cra, sheet, wb); // 左边框
                RegionUtil.setBorderRight(HSSFCellStyle.BORDER_THIN, cra, sheet, wb); // 有边框
                RegionUtil.setBorderTop(HSSFCellStyle.BORDER_THIN, cra, sheet, wb); // 上边框
            }

        }
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER); // 居中
        // 产生一行
        HSSFRow row = tool.createRow(sheet);
        HSSFCell cell;// 产生单元格
        for (Field field : fields) {
            ExcelAttribute attr = field.getAnnotation(ExcelAttribute.class);
            int col = getExcelCol(attr.column());// 获得列号
            sheet.setColumnWidth(col, 256 * attr.width());
            cell = row.createCell(col);// 创建列
            row.setHeight((short) (2 * 256));
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);// 设置列中写入内容为String类型
            String exportName;
            exportName = attr.name();
            cell.setCellValue(exportName);// 写入列名
            // 如果设置了提示信息则鼠标放上去提示.
            if (!attr.prompt().trim().equals("")) {
                setHSSFPrompt(sheet, "", attr.prompt(), 1, 100, col, col);// 这里默认设了2-101列提示.
            }
            // 如果设置了combo属性则本列只能选择不能输入
            if (attr.combo().length > 0) {
                setHSSFValidation(sheet, attr.combo(), 1, 100, col, col);// 这里默认设了2-101列只能选择不能输入.
            }
            cell.setCellStyle(style);
        }

    }

    @Override
    public void createRow(ExcelTool tool, Object t, HSSFSheet sheet, HSSFWorkbook wb) {
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框
        HSSFCell cell;// 产生单元格
        if (flag) {
            row = tool.createRow(sheet);
        }

        Field[] fields = t.getClass().getDeclaredFields();
        try {
            for (Field field : fields) {

                if (!field.isAccessible()) {
                    // 设置私有属性为可访问
                    field.setAccessible(true);
                }
                if (field.isAnnotationPresent(ExcelAttribute.class) && field.isAnnotationPresent(ExcelElement.class)) {
                    flag = false;
                    switch (ElementTypePath.getElementTypePath(field.getType().getName())) {
                        case MAP:
                            ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                            Map<?, ?> map = (Map<?, ?>) field.get(t);
                            if (map != null) {
                                StringBuilder strB = new StringBuilder();
                                for (Map.Entry<?, ?> entry : map.entrySet()) {
                                    strB.append(entry.getKey()).append(" : ").append(entry.getValue()).append(" , ");
                                }
                                if (strB.length() > 0) {
                                    strB.deleteCharAt(strB.length() - 1);
                                    strB.deleteCharAt(strB.length() - 1);
                                }

                                try {
                                    // 根据ExcelVOAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
                                    if (ea.isExport()) {
                                        cell = row.createCell(getExcelCol(ea.column()));// 创建cell
                                        cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                        cell.setCellValue(strB.toString());// 如果数据存在就填入,不存在填入空格.
                                        cell.setCellStyle(style);
                                    }
                                } catch (IllegalArgumentException e) {
                                    e.printStackTrace();
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            for (Field field : fields) {
                if (!field.isAccessible()) {
                    // 设置私有属性为可访问
                    field.setAccessible(true);
                }
                if (field.isAnnotationPresent(ExcelAttribute.class) && !field.isAnnotationPresent(ExcelElement.class)) {
                    ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                    flag = true;
                    try {
                        // 根据ExcelVOAttribute中设置情况决定是否导出,有些情况需要保持为空,希望用户填写这一列.
                        if (ea.isExport() && !ea.isImageFormat()) {
                            cell = row.createCell(getExcelCol(ea.column()));// 创建cell
                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                            cell.setCellValue(field.get(t) == null ? "" : String.valueOf(field.get(t)));// 如果数据存在就填入,不存在填入空格.
                            cell.setCellStyle(style);
                        } else if (ea.isExport() && ea.isImageFormat()) {
                            cell = row.createCell(getExcelCol(ea.column()));// 创建cell
                            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                            cell.setCellStyle(style);
                            row.setHeight((short) (30 * 40));  //有图片的行 高度应设置高点
                            createPicture(wb, sheet, row.getRowNum(), getExcelCol(ea.column()), field.get(t) == null ? "" : String.valueOf(field.get(t)));
                        }
                        //设置为超链接
                        if (ea.isLink()) {
                            //设置超链接的样式 字符颜色变成蓝色和加上下划线
                            HSSFCellStyle linkStyle = wb.createCellStyle();
                            HSSFFont cellFont = wb.createFont();
                            cellFont.setUnderline((byte) 1);
                            cellFont.setColor(HSSFColor.BLUE.index);
                            linkStyle.setFont(cellFont);
                            linkStyle.setAlignment(CellStyle.ALIGN_CENTER);//水平居中
                            linkStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);//垂直居中
                            linkStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下边框
                            linkStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);//左边框
                            linkStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);//上边框
                            linkStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);//右边框

                            cell = row.createCell(getExcelCol(ea.column()));// 创建cell
                            String value = field.get(t) == null ? "" : String.valueOf(field.get(t));
                            CreationHelper helper = wb.getCreationHelper();
                            Hyperlink hLink = helper.createHyperlink(Hyperlink.LINK_URL);
                            hLink.setAddress(value);
                            cell.setHyperlink(hLink);
                            cell.setCellStyle(linkStyle);
                            cell.setCellValue(value);
                        }
                    } catch (IllegalArgumentException | IllegalAccessException e) {
                        e.printStackTrace();
                    }
                } else if (field.isAnnotationPresent(ExcelElement.class)) {
                    flag = false;
                    switch (ElementTypePath.getElementTypePath(field.getType().getName())) {
                        case SET:
                            Set<?> set = (Set<?>) field.get(t);
                            if (set != null) {
                                for (Object object : set) {
                                    createRow(tool, object, sheet, wb);
                                }
                            }
                            break;
                        case LIST:
                            List<?> list = (List<?>) field.get(t);
                            if (list != null) {
                                for (Object object : list) {
                                    createRow(tool, object, sheet, wb);
                                }
                            }
                            break;
                        case MAP:
                            break;
                        default:
                            createRow(tool, field.get(t), sheet, wb);
                            break;
                    }
                }

            }
        } catch (SecurityException | IllegalArgumentException | IllegalAccessException e) {
            e.printStackTrace();
        }
        flag = true;
    }

    @SuppressWarnings("unchecked")
    @Override
    public List<T> dispatch(HSSFSheet sheet) {
        int rows = sheet.getPhysicalNumberOfRows();
        List<Integer> idCols = getIdCols(clazz, null);
        List<Class<?>> clazzs = getAllClass(clazz, null);
        Map<String, String> tuples = getTuple(sheet, rows);
        if (idCols.size() != clazzs.size()) {
            logger.error("class 数目不一致");
            return null;
        }
        Map<String, Map<String, Object>> instanceMap = null;
        int size = idCols.size();
        for (int i = size - 1; i > -1; i--) {
            //默认起始行为1
            for (int j = 1; j < rows; j++) {
                int childIdCol = -1;
                int idCol = idCols.get(i);
                int parentIdCol = -1;
                if (i > 0) {
                    parentIdCol = idCols.get(i - 1);
                }
                if (i < size - 1) {
                    childIdCol = idCols.get(i + 1);
                }
                instanceMap = createInstance(clazzs.get(i), j, idCol, parentIdCol, childIdCol, tuples, instanceMap);
            }

        }
        List<T> list = new ArrayList<>();
        Map<String, Object> map = instanceMap.get(idCols.get(0).toString());
        for (Map.Entry<String, Object> entry : map.entrySet()) {
            list.add((T) entry.getValue());
        }
        return list;
    }

    /**
     * 合并数据
     */
    @Override
    public int mergedRegio(Object t, HSSFSheet sheet, HSSFWorkbook wb, int rowStart) {
        //获取子节点的数目
        Field[] fields = t.getClass().getDeclaredFields();
        int rowEnd = rowStart + childNodes(t, 0) - 1;
        logger.debug(rowStart + "====>" + rowEnd);
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelAttribute.class)) {
                ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowStart, rowEnd, getExcelCol(ea.column()), getExcelCol(ea.column()));
                sheet.addMergedRegion(cellRangeAddress);
                // 使用RegionUtil类为合并后的单元格添加边框
                RegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN, cellRangeAddress, sheet, wb); // 下边框
                RegionUtil.setBorderLeft(HSSFCellStyle.BORDER_THIN, cellRangeAddress, sheet, wb); // 左边框
                RegionUtil.setBorderRight(HSSFCellStyle.BORDER_THIN, cellRangeAddress, sheet, wb); // 有边框
                RegionUtil.setBorderTop(HSSFCellStyle.BORDER_THIN, cellRangeAddress, sheet, wb); // 上边框
            } else if (field.isAnnotationPresent(ExcelElement.class) && !field.isAnnotationPresent(ExcelAttribute.class)) {
                if (!field.isAccessible()) {
                    field.setAccessible(true);
                }
                int childRowStart = rowStart;
                try {
                    switch (ElementTypePath.getElementTypePath(field.getType().getName())) {
                        case SET:
                            Set<?> set = (Set<?>) field.get(t);
                            if (set != null) {
                                for (Object object : set) {
                                    childRowStart = mergedRegio(object, sheet, wb, childRowStart);
                                }
                            }
                            break;
                        case LIST:
                            List<?> list = (List<?>) field.get(t);
                            if (list != null) {
                                for (Object object : list) {
                                    childRowStart = mergedRegio(object, sheet, wb, childRowStart);
                                }
                            }
                            break;
                        case MAP:
                            break;
                        default:
                            childRowStart = mergedRegio(field.get(t), sheet, wb, childRowStart);
                            break;
                    }
                } catch (IllegalArgumentException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        return rowEnd + 1;

    }

    /**
     * 统计子节点的叶子数目
     *
     * @param t
     * @param childNodeNum
     * @return int 子节点的数目
     * @throws
     * @Description: (一个类只支持包含一个集合)
     */
    private int childNodes(Object t, int childNodeNum) {
        Field[] fields = t.getClass().getDeclaredFields();
        boolean childNodeFlag = true;
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelElement.class) && !field.isAnnotationPresent(ExcelAttribute.class)) {
                if (!field.isAccessible()) {
                    field.setAccessible(true);
                }
                childNodeFlag = false;
                try {
                    switch (ElementTypePath.getElementTypePath(field.getType().getName())) {
                        case SET:
                            Set<?> set = (Set<?>) field.get(t);
                            if (set != null) {
                                if (set.size() == 0) {
                                    childNodeFlag = true;
                                } else {
                                    for (Object object : set) {
                                        childNodeNum = childNodes(object, childNodeNum);
                                    }
                                }

                            } else {
                                childNodeFlag = true;
                            }
                            break;
                        case LIST:
                            List<?> list = (List<?>) field.get(t);
                            if (list != null) {
                                if (list.size() == 0) {
                                    childNodeFlag = true;
                                } else {
                                    for (Object object : list) {
                                        childNodeNum = childNodes(object, childNodeNum);
                                    }
                                }
                            } else {
                                childNodeFlag = true;
                            }
                            break;
                        case MAP:
                            break;
                        default:
                            childNodeNum = childNodes(field.get(t), childNodeNum);
                            break;
                    }
                } catch (IllegalArgumentException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        if (childNodeFlag) {
            childNodeNum++;
        }
        return childNodeNum;
    }

    /**
     * 创建实例对象 并装配关系
     *
     * @param clazz       当前实例对象的类
     * @param row         当前实例对象的行
     * @param idCol       当前实例对象的id列
     * @param parentIdCol 父实例对象的id列
     * @param tuples      存放所有的值
     * @param instanceMap 装配后的对象
     * @throws
     */
    @SuppressWarnings("unchecked")
    private Map<String, Map<String, Object>> createInstance(Class<?> clazz, int row, int idCol, int parentIdCol, int childIdCol, Map<String, String> tuples, Map<String, Map<String, Object>> instanceMap) {
        if (!tuples.containsKey(row + "," + idCol)) {
            return instanceMap;
        }
        // TODO判断 是否存在该实例
        if (instanceMap == null) {
            instanceMap = new HashMap<>();
        }
        Object entity = null;
        try {
            entity = clazz.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
        //同一类对象的map 以 id_parentId作为键
        Map<String, Object> colMap;
        if (instanceMap.containsKey(idCol + "")) {
            colMap = instanceMap.get(idCol + "");
        } else {
            colMap = new HashMap<>();
            instanceMap.put(idCol + "", colMap);
        }
        //判断是否有父级对象
        String id = tuples.get(row + "," + idCol);
        if (parentIdCol > -1) {
            String parentId = tuples.get(row + "," + parentIdCol);
            colMap.put(parentId + "_" + id, entity);

        } else {
            colMap.put(id, entity);
        }
//		logger.debug("colMap===========>"+colMap.size());
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (!field.isAccessible()) {
                field.setAccessible(true);
            }
            //对普通属性进行设值
            if (field.isAnnotationPresent(ExcelAttribute.class) && !field.isAnnotationPresent(ExcelElement.class)) {
                ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                String value = tuples.get(row + "," + getExcelCol(ea.column()));
                Class<?> fieldType = field.getType();
                try {
                    if (String.class == fieldType) {
                        field.set(entity, String.valueOf(value));
                    } else if ((Integer.TYPE == fieldType) || (Integer.class == fieldType)) {
                        field.set(entity, Integer.parseInt(value));
                    } else if ((Long.TYPE == fieldType) || (Long.class == fieldType)) {
                        field.set(entity, Long.valueOf(value));
                    } else if ((Float.TYPE == fieldType) || (Float.class == fieldType)) {
                        field.set(entity, Float.valueOf(value));
                    } else if ((Short.TYPE == fieldType) || (Short.class == fieldType)) {
                        field.set(entity, Short.valueOf(value));
                    } else if ((Double.TYPE == fieldType) || (Double.class == fieldType)) {
                        field.set(entity, Double.valueOf(value));
                    } else if (Character.TYPE == fieldType) {
                        if ((value != null) && (value.length() > 0)) {
                            field.set(entity, Character.valueOf(value.charAt(0)));
                        }
                    }
                } catch (NumberFormatException e) {
                    e.printStackTrace();
                } catch (IllegalArgumentException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            } else if (field.isAnnotationPresent(ExcelElement.class)) {
                Map<String, Object> map = instanceMap.get(childIdCol + "");
                List<Object> entitys = new ArrayList<>();
                if (map != null) {
                    for (Map.Entry<String, Object> entry : map.entrySet()) {
                        String key = entry.getKey();
                        if (key.indexOf("_") > -1) {
                            String[] str = key.split("_");
                            String childParentId = str[0];
                            if (childParentId.equals(id)) {
                                entitys.add(entry.getValue());
                            }
                        }
                    }
                }

                // 进行类注入
                String typeName = field.getType().getName();
                try {

//					logger.debug("entitys:"+entitys.size());
                    switch (ElementTypePath.getElementTypePath(typeName)) {
                        case SET:
                            Set<Object> set = (Set<Object>) field.get(entity);
                            if (set == null) {
                                set = new HashSet<>();
                                field.set(entity, set);
                            }
                            for (Object object : entitys) {
                                set.add(object);
                            }
                            break;
                        case LIST:
                            List<Object> list = (List<Object>) field.get(entity);
                            if (list == null) {
                                list = entitys;
                                field.set(entity, list);
                            } else {
                                for (Object object : entitys) {
                                    list.add(object);
                                }
                            }
                            break;
                        case MAP:
                            if (field.isAnnotationPresent(ExcelAttribute.class)) {
                                if (getClass(field.getGenericType(), 0).getName().equals("java.lang.String")) {
                                    Map<String, String> imap = (Map<String, String>) field.get(entity);
                                    if (imap == null) {
                                        imap = new HashMap<>();
                                        field.set(entity, imap);
                                    }
                                    ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                                    String value = tuples.get(row + "," + getExcelCol(ea.column()));
                                    if (value.contains(",")) {
                                        String[] str = value.split(",");
                                        for (String string : str) {
                                            if (string.contains(":")) {
                                                String[] keyAndVlaue = string.split(":");
                                                if (keyAndVlaue.length == 2) {
                                                    imap.put(keyAndVlaue[0].trim(), keyAndVlaue[1].trim());
                                                }
                                            }
                                        }
                                    }
                                }


                            }
                            break;
                        default:
                            if (entitys.size() == 1) {
                                field.set(entity, entitys.get(0));
                            }
                            break;

                    }
                } catch (IllegalArgumentException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
        return instanceMap;
    }


    /**
     * 获取所有标识列（用来获取父级id）
     *
     * @param clazz
     * @param idCols
     * @return List<Integer> 返回类型
     * @throws
     */
    private List<Integer> getIdCols(Class<?> clazz, List<Integer> idCols) {
        if (idCols == null) {
            idCols = new ArrayList<>();
        }
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelID.class) && field.isAnnotationPresent(ExcelAttribute.class)) {
                ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                idCols.add(getExcelCol(ea.column()));
            }
        }
        for (Field field : fields) {
            // 此处集合需要做判断
            if (field.isAnnotationPresent(ExcelElement.class)) {
                clazz = getClass(field.getGenericType(), 0);
//				logger.debug(clazz);
                getIdCols(clazz, idCols);
            }
        }
        return idCols;
    }

    private List<Class<?>> getAllClass(Class<?> clazz, List<Class<?>> clazzs) {
        if (clazzs == null) {
            clazzs = new ArrayList<>();
        }
        clazzs.add(clazz);
        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            // 此处集合需要做判断
            if (field.isAnnotationPresent(ExcelElement.class) && !field.isAnnotationPresent(ExcelAttribute.class)) {
                clazz = getClass(field.getGenericType(), 0);
                getAllClass(clazz, clazzs);
            }
        }
        return clazzs;
    }

    /**
     * 获取 excel中的数据 元组
     *
     * @param sheet
     * @return List<List   <   String>> excel元组
     * @throws
     */
    private Map<String, String> getTuple(HSSFSheet sheet, int rows) {

        Map<String, String> tuples = new HashMap<>();
        //获取列
        List<Field> fields = getAllField(clazz, null);

        // 从第2行开始取数据,默认第一行是表头.
        for (int i = 1; i < rows; i++) {

            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelAttribute.class)) {
                    ExcelAttribute ea = field.getAnnotation(ExcelAttribute.class);
                    int col = getExcelCol(ea.column());
                    if (ExcelTool.isMergedRegion(sheet, i, col)) {
                        // 以(行,列）作为一个元组
                        String key = i + "," + col;
//						logger.debug(key);
                        String value = ExcelTool.getMergedRegionValue(sheet, i, col);
//						logger.debug(key+" ===>"+value);  
                        tuples.put(key, value);
                    }
                }
            }

        }
        return tuples;
    }


    public static void createPicture(HSSFWorkbook workbook, HSSFSheet sheet, int row, int column, String filePaths) {
        try {
            BufferedImage bufferImg;
            String[] fileArray = filePaths.split(",");
            //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
            HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
            sheet.setColumnWidth(0, 256 * 20);
            int x = 1024 / fileArray.length;
            int dax1 = 0;
            int dax2 = x - 1;
            for (String filePath : fileArray) {
                ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
                bufferImg = ImageIO.read(new File(filePath));
                ImageIO.write(bufferImg, "jpg", byteArrayOut);
                HSSFClientAnchor anchor = new HSSFClientAnchor(dax1, 0, dax2, 255, (short) column, row, (short) column, row);
                //Move With Cells but Do Not Resize 随单元格移动但不调整图片大小
                anchor.setAnchorType(2);
                //插入图片
                patriarch.createPicture(anchor, workbook.addPicture(byteArrayOut.toByteArray(), HSSFWorkbook.PICTURE_TYPE_JPEG));
                //为第二个图片的增加x轴偏移量
                dax1 += x;
                dax2 += x;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}


