package com.xcb.pojo;

import java.io.Serializable;

/**
 * 
 * @类功能说明: 封装的excel节点数据类
 * @类修改者:
 * @修改日期:
 * @修改说明:
 * @作者: Aaron
 * @创建时间: 2013年10月16日 下午8:14:38
 * @版本: 1.0.0
 */
public class ExcelNode implements Serializable {

	/**
	 * @类名: ExcelNode.java
	 * @描述: TODO
	 * 
	 */

	private static final long serialVersionUID = 3391962399800252112L;

	/**
	 * 暂定两个节点
	 */
	private String noteName1;
	private String noteName2;

	public String getNoteName1() {
		return noteName1;
	}

	public void setNoteName1(String noteName1) {
		this.noteName1 = noteName1;
	}

	public String getNoteName2() {
		return noteName2;
	}

	public void setNoteName2(String noteName2) {
		this.noteName2 = noteName2;
	}

	public String toString() {
		// TODO Auto-generated method stub

		StringBuilder sb = new StringBuilder(50);

		sb.append("Excel节点如下：");
		sb.append("noteName1:");
		sb.append(noteName1);
		sb.append("noteName2:");
		sb.append(noteName2);
		return sb.toString();
	}

}
