package com.picc.nf.util.poiUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 一个通用的将List<T>中数据导出为Excel文档的工具类
 * 
 * @author 
 */
public class ExcelExporter {
	
	private static final Log log = LogFactory.getLog(ExcelExporter.class);
	
	/**
	 * 弃用方法
     * excel文件导入
     * 该方法在传统方法加以修改，支持不同版本导入，最终返回一个list集合，集合中封装的是object数组，然后根据业务需要封装自己想要的pojo对象。
     * 该方法支持指定服务端文件路径上传和客户端上传，如果是服务端指定路径上传则file赋值null，反之，则filePaht赋值null
     * 前端页面代码以下4行
     * <form action="upload" enctype="multipart/form-data">
     * <input type="file" name="myFile" />
     * <input type="submit" value="Upload! " />
     * </form>
     * springMvc action 中使用以下代码，以下两行
     * MultipartHttpServletRequest mulRequest = (MultipartHttpServletRequest) request;
     * MultipartFile file = mulRequest.getFile("excel");
     * @param file
     * @param filePath "d:/test.xls";
     * @return
     */
    public static List excelToList(InputStream input,String filePath){
        List list=null;//返回对象
       
        try { 
            Workbook workBook = null;
            //以下做法是为了区分不同版本，然后使用XSSF或HSSF
            String flag="";//XSSF或HSSF的标记
            try {
                workBook = new XSSFWorkbook(input);
                flag="XSSF";
            } catch (Exception ex) {
                flag="HSSF";
            }
            
            //根据标记使用不同的方法来解析excel
            if("XSSF".equals(flag)){
                 XSSFSheet sheet = (XSSFSheet) workBook.getSheetAt(0);  
                 if (sheet != null)  
                 {      //list
                      list = new ArrayList();    
                     for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++)  
                     {   //object
                          XSSFRow row = sheet.getRow(i);
                           Object[] obj = new Object[row.getPhysicalNumberOfCells()];
                         for (int j = 0; j < row.getPhysicalNumberOfCells(); j++)  
                         {   //column
                              XSSFCell cell = row.getCell(j);  
                             String cellStr = cell.toString();  
                             log.info("【"+cellStr+"】 ");
                             obj[j] = cellStr;
                         }  
                         System.out.println();
                         list.add(obj);
                     }
                     log.info("当前excel总条数"+list.size()); 
                 }  
            }else if ("HSSF".equals(flag)){
                 HSSFSheet sheet = (HSSFSheet) workBook.getSheetAt(0);  
                 if (sheet != null)  
                 {      //list
                      list = new ArrayList();    
                     for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++)  
                     {   //object
                          HSSFRow row = sheet.getRow(i);
                          Object[] obj = new Object[row.getPhysicalNumberOfCells()];
                         for (int j = 0; j < row.getPhysicalNumberOfCells(); j++)  
                         {   //column
                             HSSFCell cell = row.getCell(j);  
                             String cellStr = cell.toString();  
                             log.info("【"+cellStr+"】 ");
                             obj[j] = cellStr;
                         }  
                         System.out.println();
                         list.add(obj);
                     }
                     log.info("当前excel总条数"+list.size()); 
                 }  
            }else{
                log.error("导入excel,发生未知错误！");
            }
        }  
        catch (Exception e) { 
            e.printStackTrace();  
        }
        return list;
    }

	/**
	 * 根据ExcelEntity等参数生成Workbook
	 * 
	 * @param entity
	 * @return
	 * @throws Exception
	 */
	public static <T> Workbook export2Excel(ExcelEntity<T> entity) throws Exception {
		Workbook workbook = export2Excel(entity.getHeader(), entity.getFooter(), entity.getSheetName(),
				entity.getColumnNames(), entity.getMethodNames(), entity.getEntities());
		return workbook;
	}

	/**
	 * 根据给定参数导出Excel文档
	 * 
	 * @param headerTitle
	 *            题头
	 * @param footer
	 *            脚注
	 * @param sheetName
	 * @param columnNames
	 *            表头名称
	 * @param methodNames
	 * @param entities
	 * @return
	 * @throws Exception
	 */
	public static <T> Workbook export2Excel(String headerTitle, String footerTitle, String sheetName,
			String[] columnNames, String[] methodNames, List<T> entities) throws Exception {
		if (methodNames.length != columnNames.length)
			throw new IllegalArgumentException("methodNames.length should be equal to columnNames.length:"
					+ columnNames.length + " " + methodNames.length);
		Workbook newWorkBook2007 = new XSSFWorkbook();
		Sheet sheet = newWorkBook2007.createSheet(sheetName);

		// 设置题头
		Header header = sheet.getHeader();
		header.setCenter(headerTitle);
		// 设置脚注
		Footer footer = sheet.getFooter();
		footer.setCenter(footerTitle);

		int[] columnWidths = new int[columnNames.length];
		// 创建表头
		createTableHeader(sheet, 0, headerTitle, columnNames, columnWidths);
		// 填充表内容
		createTableContent(sheet, 1, methodNames, columnWidths, entities);

		return newWorkBook2007;

	}

	/**
	 * 创建表头
	 * 
	 * @param sheet
	 * @param index
	 *            表头开始的行数
	 * @param headerTitle
	 *            题头
	 * @param columnNames
	 * @param columnWidths
	 */
	private static void createTableHeader(Sheet sheet, int index, String headerTitle, String[] columnNames,
			int[] columnWidths) {

		Row headerRow = sheet.createRow(index);

		/* 格式设置 */
		// 设置字体
		Font font = sheet.getWorkbook().createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);// 粗体显示
		// 设置背景色
		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		style.setFont(font);

		for (int i = 0; i < columnNames.length; i++) {
			Cell headerCell = headerRow.createCell(i);
			headerCell.setCellStyle(style);
			headerCell.setCellValue(columnNames[i]);
		}

		for (int i = 0; i < columnNames.length; i++) {
			columnWidths[i] = (columnNames[i].getBytes().length + 2) * 256;
			sheet.setColumnWidth(i, columnWidths[i]);
		}

	}

	/**
	 * 创建表格内容
	 * 
	 * @param sheet
	 * @param rowIndexBegin
	 *            表内容开始的行数
	 * @param methodNames
	 *            T对象的方法名
	 * @param columnWidths
	 * @param entities
	 * @throws Exception
	 */
	private static <T> void createTableContent(Sheet sheet, int rowIndexBegin, String[] methodNames, int[] columnWidths,
			List<T> entities) throws Exception {
		Class<? extends Object> clazz = null;
		if (entities.size() > 0)
			clazz = entities.get(0).getClass();

		String content = null;
		for (T t : entities) {
			Row row = sheet.createRow(rowIndexBegin++);
			for (int i = 0; i < methodNames.length; i++) {
				Cell cell = row.createCell(i);
				Method method = clazz.getMethod(methodNames[i], null);
				Object object = method.invoke(t, null);
				object = object == null ? "" : object;
				if (object.getClass().equals(Date.class)) {// 对日期格式进行特殊处理
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					content = sdf.format((Date) object);
					cell.setCellValue(content);
				} else {
					content = object.toString();
					cell.setCellValue(content);
				}
				int columnWidth = (content.getBytes().length + 2) * 256;
				if (columnWidth > columnWidths[i]) {// 如果实际内容宽度大于对应的表头宽度，则设置为实际内容宽度
					columnWidths[i] = columnWidth;
					sheet.setColumnWidth(i, columnWidths[i]);
				}

			}
		}
	}

	public static <T> void testPOI(String[] columnNames, String[] methodNames, List<T> entities) throws Exception {
		String sheetName = "Test";
		String title = "标题栏";
		String dstFile = "d:/temp/test.xlsx";
		Workbook newWorkBook2007 = new XSSFWorkbook();
		Sheet sheet = newWorkBook2007.createSheet(sheetName);
		int[] columnWidths = new int[columnNames.length];
		// 创建表头
		createTableHeader(sheet, 0, title, columnNames, columnWidths);
		// 填充表内容
		createTableContent(sheet, 1, methodNames, columnWidths, entities);
		// 保存为文件
		saveWorkBook2007(newWorkBook2007, dstFile);
		System.out.println("end");

	}

	/**
	 * 将workbook2007存为文件
	 * 
	 * @param workbook2007
	 * @param dstFile
	 */
	public static void saveWorkBook2007(Workbook workbook2007, String dstFile) {
		File file = new File(dstFile);
		OutputStream os = null;
		try {
			os = new FileOutputStream(file);
			workbook2007.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (os != null) {
				try {
					os.close();
				} catch (IOException e) {
					System.out.println(e);
				}
			}
		}
	}

	/**
	 * 测试方法
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		// 准备数据
//		List<Wind> winds = new ArrayList<>();// Wind有三个方法:getLocation、getSpeed、getTimestamp
//		for (int i = 0; i < 10; i++) {
//			Wind wind = new Wind();
//			wind.setLocation(i);
//			wind.setSpeed(i * 10);
//			wind.setTimestamp("2016/3/2" + i);
//			winds.add(wind);
//		}
//		String[] columnNames = { "地点", "速度", "时间", "环境" };
//		String[] methodNames = { "getLocation", "getSpeed", "getTimestamp", "getSurroundings" };
//		// String fileName = "d:/temp/excel1.xlsx";
//		String fileName = "d:/excel1.xlsx";
//		// 生成ExcelEntity实体，包含4个必备参数
//		ExcelEntity<Wind> excelEntity = new ExcelEntity<>(fileName, columnNames, methodNames, winds);
//		// excelEntity.setHeader("题头");
//		// excelEntity.setFooter("脚注");
//		Workbook excel = ExcelExporter.export2Excel(excelEntity);
//		// ExcelExporter.export2Excel("题头","脚注", "sheet1", columnNames,
//		// methodNames, winds);//也可以这样调用,无需新建ExcelEntity对象
//		// 将Workbook存为文件
//		ExcelExporter.saveWorkBook2007(excel, excelEntity.getFileName());

		System.out.println("导出完成！");

	}

}