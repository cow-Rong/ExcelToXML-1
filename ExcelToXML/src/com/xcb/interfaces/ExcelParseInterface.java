package com.xcb.interfaces;

import java.util.List;

import com.xcb.pojo.ExcelNode;

/**
 * 
* @类功能说明: 解析excel的公共接口
 * @类修改者:   
 * @修改日期:   
 * @修改说明:   
 * @作者:      Aaron
 * @创建时间:  2013年10月16日 下午8:11:25
 * @版本:      1.0.0
 */
public interface ExcelParseInterface {

	public List<ExcelNode> parse(String path);
}
