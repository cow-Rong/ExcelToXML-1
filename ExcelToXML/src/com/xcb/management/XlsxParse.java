package com.xcb.management;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.xcb.interfaces.ExcelParseInterface;
import com.xcb.pojo.ExcelNode;

/**
 * 
 * @类功能说明: 解析XLSX类型的excel
 * @类修改者:
 * @修改日期:
 * @修改说明:
 * @作者: Aaron
 * @创建时间: 2013年10月16日 下午8:25:34
 * @版本: 1.0.0
 */
public class XlsxParse implements ExcelParseInterface {

	private XSSFWorkbook xssfWorkbook = null;
	private XSSFSheet xssfSheet = null;
	private XSSFRow xssfRow = null;
	private XSSFCell xssfCell = null;

	public List<ExcelNode> parse(String path) {
		// TODO Auto-generated method stub

		try {
			xssfWorkbook = new XSSFWorkbook(path);

			/**
			 * 循环工作表
			 */
			for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {

				xssfSheet = xssfWorkbook.getSheetAt(numSheet);

				if (null == xssfSheet) {

					continue;
				}

				/**
				 * 循环遍历行
				 */
				for (int rowNum = 0; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {

					xssfRow = xssfSheet.getRow(rowNum);
					if (null == xssfRow) {

						continue;
					}

					/**
					 * 循环列
					 */

					for (int cellNum = 0; cellNum <= xssfRow.getLastCellNum(); cellNum++) {

						xssfCell = xssfRow.getCell(cellNum);

						if (null == xssfCell) {

							continue;
						}

						System.out.println("   " + getValue(xssfCell));
					}

				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return null;
	}

	private String getValue(XSSFCell xssfCell) {

		if (xssfCell.getCellType() == xssfCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(xssfCell.getBooleanCellValue());
		} else if (xssfCell.getCellType() == xssfCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(xssfCell.getNumericCellValue());
		} else {
			return String.valueOf(xssfCell.getStringCellValue());
		}

	}

}
