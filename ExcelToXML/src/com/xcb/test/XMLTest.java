package com.xcb.test;

import com.xcb.management.XlsParse;
import com.xcb.management.XlsxParse;

public class XMLTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		/*XlsParse xlsParse = new XlsParse();
		
		xlsParse.parse("d:\\a.xls");*/
		
		XlsxParse xlsxParse = new XlsxParse();
		
		xlsxParse.parse("d:\\b.xlsx");
		
	}

}
