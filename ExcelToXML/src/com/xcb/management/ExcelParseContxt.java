package com.xcb.management;

import java.util.List;

import com.xcb.interfaces.ExcelParseInterface;
import com.xcb.pojo.ExcelNode;

public class ExcelParseContxt {

	/**
	 * 解析excel接口对象
	 */
	private ExcelParseInterface excelParseInterface;

	public ExcelParseContxt(ExcelParseInterface excelParseInterface) {

		this.excelParseInterface = excelParseInterface;
	}

	public List<ExcelNode> parse(String path) {

		return excelParseInterface.parse(path);
	}
}
