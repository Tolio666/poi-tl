package com.deepoove.poi.util;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.List;

/**
 * 样式复制
 *
 * @author 奔波儿灞
 * @since 1.6.0
 */
public final class PoiStyleCopyUtils {

	private PoiStyleCopyUtils() {
		throw new IllegalStateException("Utils");
	}

	/**
	 * 复制一行的样式
	 *
	 * @param sourceRow 样式来源
	 * @param targetRow 目标行
	 */
	public static void copyRowStyle(XWPFTableRow sourceRow, XWPFTableRow targetRow) {
		// 复制行属性
		targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
		List<XWPFTableCell> cellList = sourceRow.getTableCells();
		if (null == cellList) {
			return;
		}
		// 复制列及其属性和内容
		XWPFTableCell targetCell;
		for (XWPFTableCell sourceCell : cellList) {
			targetCell = targetRow.addNewTableCell();
			// 列属性
			targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
			// 段落属性
			targetCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
		}
	}

}
