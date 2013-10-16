package com.xcb.management;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.xcb.interfaces.ExcelParseInterface;
import com.xcb.pojo.ExcelNode;

/**
 * 
 * @类功能说明: 解析 XLS类型的excel
 * @类修改者:
 * @修改日期:
 * @修改说明:
 * @作者: Aaron
 * @创建时间: 2013年10月16日 下午8:25:03
 * @版本: 1.0.0
 */
public class XlsParse implements ExcelParseInterface {

	private HSSFWorkbook hssfWorkbook = null;
	private InputStream is = null;
	private HSSFSheet hssfSheet = null;

	private HSSFRow hssfRow = null;
	private HSSFCell hssfCell = null;

	public List<ExcelNode> parse(String path) {
		// TODO Auto-generated method stub

		try {
			/**
			 * 创建一个文件流
			 */
			is = new FileInputStream(path);
			/**
			 * 创建一个wokebook
			 */
			hssfWorkbook = new HSSFWorkbook(is);
			/**
			 * 循环工作簿
			 */
			for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {

				/**
				 * 获取sheet实例
				 */
				hssfSheet = hssfWorkbook.getSheetAt(numSheet);

				if (null == hssfSheet) {

					continue;
				}

				/**
				 * 循环行Row
				 */

				for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {

					hssfRow = hssfSheet.getRow(rowNum);

					if (null == hssfRow) {

						continue;
					}

					/**
					 * 循环列
					 */
					for (int cellNum = 0; cellNum <= hssfRow.getLastCellNum(); cellNum++) {

						hssfCell = hssfRow.getCell(cellNum);

						if (null == hssfCell) {

							continue;
						}

						System.out.println("   " + getValue(hssfCell));
					}
				}
			}

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return null;
	}

	private String getValue(HSSFCell hssfCell) {

		if (hssfCell.CELL_TYPE_BOOLEAN == hssfCell.getCellType()) {

			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.CELL_TYPE_NUMERIC == hssfCell.getCellType()) {

			return String.valueOf(hssfCell.getNumericCellValue());
		} else {

			return String.valueOf(hssfCell.getStringCellValue());
		}
	}

}
