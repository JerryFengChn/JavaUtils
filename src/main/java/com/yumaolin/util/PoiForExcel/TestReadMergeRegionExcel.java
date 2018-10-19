package com.yumaolin.util.PoiForExcel;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;


/**
 * 
 * @author wcyong
 * 
 * @date 2013-6-21
 */
public class TestReadMergeRegionExcel {

	@Test
	public void testReadExcel() throws IOException {
		readExcelToObj("d:\\套餐配置模板.xlsx");
	}

	/**
	 * 读取excel数据
	 * 
	 * @param path
	 * @throws IOException
	 */
	private void readExcelToObj(String path) throws IOException {
		InputStream input = new FileInputStream(new File(path));
		BufferedInputStream bufferedInput = (BufferedInputStream) FileMagic.prepareToCheckMagic(input);
		Workbook wb = null;
		if (FileMagic.valueOf(bufferedInput) == FileMagic.OOXML) {
			wb = new XSSFWorkbook(bufferedInput);
		} else if (FileMagic.valueOf(bufferedInput) == FileMagic.OLE2) {
			wb = new HSSFWorkbook(bufferedInput);
		}
		readExcel(wb, 0, 1, 0);
		IOUtils.closeQuietly(bufferedInput);
	}

	/**
	 * 读取excel文件
	 * 
	 * @param wb
	 * @param sheetIndex sheet页下标：从0开始
	 * @param startReadLine 开始读取的行:从0开始
	 * @param tailLine 去除最后读取的行
	 */
	private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) {
		Sheet sheet = wb.getSheetAt(sheetIndex);
		Row row = null;

		for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
			row = sheet.getRow(i);
			for (Cell c : row) {
				boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
				// 判断是否具有合并单元格
				if (isMerge) {
					String rs = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
					System.out.print(rs + "|");
				} else {
					System.out.print(getCellValue(c) + "|");
				}
			}
			System.out.println();
		}

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
					return getCellValue(fCell);
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

	/**
	 * 获取单元格的值
	 * 
	 * @param cell
	 * @return
	 */
	public String getCellValue(Cell cell) {
		String cellvalue = "";
		if (cell != null) {
			// 判断当前Cell的Type
			switch (cell.getCellTypeEnum()) {
				// 如果当前Cell的Type为NUMERIC
				case NUMERIC:
					BigDecimal db = new BigDecimal(String.valueOf(cell.getNumericCellValue()));// 避免精度问题，先转成字符串
					cellvalue = db.toPlainString();
					break;
				case FORMULA: {
					// 判断当前的cell是否为Date
					if (DateUtil.isCellDateFormatted(cell)) {
						Date date = cell.getDateCellValue();
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
						cellvalue = sdf.format(date);
					} else {// 如果是纯数字
						DecimalFormat df = new DecimalFormat("0");
						cellvalue = df.format(cell.getNumericCellValue());// 取得当前Cell的数值
					}
					break;
				}
				case STRING: {// 如果当前Cell的Type为STRIN
					// 取得当前的Cell字符串
					cellvalue = cell.getRichStringCellValue().getString();
					break;
				}
				default: {// 默认的Cell值
					cellvalue = "";
				}
			}
		} else {
			cellvalue = "";
		}
		return cellvalue.trim();
	}
}