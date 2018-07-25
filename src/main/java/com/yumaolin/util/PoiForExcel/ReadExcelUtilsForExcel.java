package com.yumaolin.util.PoiForExcel;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelUtilsForExcel {
	private Sheet sheet;
	private CellStyle setBorder;
	private Font font;
	
	public Workbook createWorkBook(String path) throws IOException{
		InputStream input = new FileInputStream(new File(path));
		BufferedInputStream bufferedInput = (BufferedInputStream) FileMagic.prepareToCheckMagic(input);
		Workbook wb = null;
		if(FileMagic.valueOf(bufferedInput) == FileMagic.OOXML){
			wb = new XSSFWorkbook(bufferedInput);
		}else if(FileMagic.valueOf(bufferedInput) == FileMagic.OLE2){
			wb = new HSSFWorkbook(bufferedInput);
		}else{
			throw new IOException("该文件类型不是excel文件类型!");
		}
		IOUtils.closeQuietly(bufferedInput);
		return wb;
	}

	/**
	 * 读取excel文件
	 * @param wb
	 * @param sheetIndex sheet页下标：从0开始
	 * @param startReadLine 开始读取的行:从0开始
	 * @param tailLine 去除最后读取的行
	 */
	public String[] readExcelContent(Workbook wb, int sheetIndex, int startReadLine, int tailLine){
		Sheet sheet = wb.getSheetAt(sheetIndex);
		// 得到总行数
		int rowNum = sheet.getLastRowNum()-tailLine;
		String[] content = new String[rowNum-startReadLine];
		// 正文内容应该从第二行开始,第一行为表头的标题
		Row row = null;
		int emptyCount = 0;//计算空行的个数
		for (int i = startReadLine;i<rowNum;i++) {
			row = sheet.getRow(i);
			if(row == null){
				continue;
			}
			StringBuilder cellValue = new StringBuilder(30);
			for (Cell c : row) {
				boolean isMerge = isMergedRegion(sheet,i,c.getColumnIndex());
				// 判断是否具有合并单元格
				if (isMerge) {
					String value = getMergedRegionValue(sheet,row.getRowNum(),c.getColumnIndex());
					if(StringUtils.isNotBlank(value)){
						cellValue.append(value).append("|");
					}
				} else {
					String value = getCellFormatValue(c);
					if(StringUtils.isNotBlank(value)){
						cellValue.append(value).append("|");
					}
				}
			}
			String value = cellValue.toString();
			if(StringUtils.isNotEmpty(value)){
				content[i-startReadLine-emptyCount] = value;
			}else{
				++emptyCount;
			}
		}
		content = Arrays.copyOf(content,content.length-emptyCount);
		setBorder = wb.createCellStyle();
		font = wb.createFont();
		return content;
	}
	
	/**
	 * 判断指定的单元格是否是合并单元格
	 * 
	 * @param sheet
	 * @param row 行下标
	 * @param column 列下标
	 * @return
	 */
	private boolean isMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}
	
	/**
	 * 获取合并单元格的值
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public String getMergedRegionValue(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					Row fRow = sheet.getRow(firstRow);
					Cell fCell = fRow.getCell(firstColumn);
					return getCellFormatValue(fCell);
				}
			}
		}
		return null;
	}
	
	/**
	 * 判断合并了行
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	private boolean isMergedRow(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row == firstRow && row == lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}
	
	/**
	 * 判断sheet页中是否含有合并单元格
	 * 
	 * @param sheet
	 * @return
	 */
	private boolean hasMerged(Sheet sheet) {
		return sheet.getNumMergedRegions() > 0 ? true : false;
	}
	
	/**
	 * 合并单元格
	 * 
	 * @param sheet
	 * @param firstRow 开始行
	 * @param lastRow 结束行
	 * @param firstCol 开始列
	 * @param lastCol 结束列
	 */
	private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	private static String getCellFormatValue(Cell cell) {
		String cellvalue = "";
		if (cell != null) {
			// 判断当前Cell的Type
			switch (cell.getCellTypeEnum()){
			// 如果当前Cell的Type为NUMERIC
			case NUMERIC:
				BigDecimal db = new BigDecimal(String.valueOf(cell.getNumericCellValue()));// 避免精度问题，先转成字符串
				cellvalue = db.toPlainString();
				break;
			case FORMULA:{
				// 判断当前的cell是否为Date
				if (DateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					cellvalue = sdf.format(date);
				}else{// 如果是纯数字
					DecimalFormat df = new DecimalFormat("0");
					cellvalue = df.format(cell.getNumericCellValue());//取得当前Cell的数值
				}
				break;
			}
			case STRING:{// 如果当前Cell的Type为STRIN
				// 取得当前的Cell字符串
				cellvalue = cell.getRichStringCellValue().getString();
				break;
			}default:{// 默认的Cell值
				cellvalue = "";
			}
			}
		} else {
			cellvalue = "";
		}
		return cellvalue.trim();
	}

	public void writeInTemplate(String newContent, int beginRow,int beginCell, boolean flag){
		Row row = sheet.getRow(beginRow);
		if (null == row) {
			// 如果不做空判断，你必须让你的模板文件画好边框，beginRow和beginCell必须在边框最大值以内
			// 否则会出现空指针异常
			row = sheet.createRow(beginRow);
		}
		Cell cell = row.getCell(beginCell);
		sheet.autoSizeColumn(beginCell);
		if (null == cell) {
			cell = row.createCell(beginCell);
		}
		// 设置存入内容为字符串
		//cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellType(CellType.STRING);
		getHssfCellStyle(flag);
		cell.setCellStyle(setBorder);
		// 向单元格中放入值
		cell.setCellValue(newContent);
	}

	public void getHssfCellStyle(boolean flag) {
		// cell.setCellStyle(styleFactory.getHeaderStyle());
		font.setFontHeightInPoints((short) 12); // 字体高度
		//font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		font.setBold(true);
		font.setColor(HSSFColorPredefined.BLACK.getIndex());
		/*setBorder.setBorderBottom(CellStyle.BORDER_THIN);
		setBorder.setBorderLeft(CellStyle.BORDER_THIN);
		setBorder.setBorderTop(CellStyle.BORDER_THIN);
		setBorder.setBorderRight(CellStyle.BORDER_THIN);*/
		setBorder.setBorderBottom(BorderStyle.THIN);
		setBorder.setBorderLeft(BorderStyle.THIN);
		setBorder.setBorderTop(BorderStyle.THIN);
		setBorder.setBorderRight(BorderStyle.THIN);
		setBorder.setFont(font);
	}
	
	public static void main(String[] args) throws Exception {
		/**
		 * 如果是xlsx文件，则按照xlsx文件类型读取
		 */
		ReadExcelUtilsForExcel readExcel = new ReadExcelUtilsForExcel();
		Workbook wb = readExcel.createWorkBook("d:\\套餐配置模板.xlsx");
		String[] list = readExcel.readExcelContent(wb,0,1,0);
		for(String str : list){
			System.out.println(str);
		}
	}
}
