package com.deepoove.poi.policy;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PatternCell;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import com.deepoove.poi.util.PoiStyleCopyUtils;
import com.deepoove.poi.util.TableTools;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.standard.SpelExpressionParser;

import java.util.Collections;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * 表格模式匹配RenderPolicy
 * table|5:var
 *
 * @author 奔波儿灞
 * @since 1.6.0
 */
public class TablePatternRenderPolicy extends AbstractPatternRenderPolicy {

	public static final String NAME = "table";

	private static final int ROW_LIMIT = 2;

	private static final Pattern PATTERN = Pattern.compile("\\{(.+)}");

	private final ExpressionParser parser = new SpelExpressionParser();

	private int generateLength;

	@Override
	public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
		String attr = getAttr();
		try {
			generateLength = Integer.valueOf(attr);
		} catch (NumberFormatException e) {
			throw new RenderException("Table render attr must be number: " + attr);
		}

		NiceXWPFDocument doc = template.getXWPFDocument();
		RunTemplate runTemplate = (RunTemplate) eleTemplate;
		XWPFRun run = runTemplate.getRun();
		run.setText("", 0);
		try {
			if (!TableTools.isInsideTable(run)) {
				throw new IllegalStateException("The template tag " + runTemplate.getSource() + " must be inside a table");
			}
			// w:tbl-w:tr-w:tc-w:p-w:tr
			XmlCursor newCursor = ((XWPFParagraph) run.getParent()).getCTP().newCursor();
			newCursor.toParent();
			newCursor.toParent();
			newCursor.toParent();
			XmlObject object = newCursor.getObject();
			XWPFTable table = doc.getTableByCTTbl((CTTbl) object);
			doRender(table, data);
		} catch (Exception e) {
			throw new RenderException("dynamic table error:" + e.getMessage(), e);
		}
	}

	private void doRender(XWPFTable table, Object data) {
		List<XWPFTableRow> rows = table.getRows();
		if (rows.size() < ROW_LIMIT) {
			throw new RenderException("The render table must be " + ROW_LIMIT + " rows");
		}
		XWPFTableRow row = rows.get(1);
		List<PatternCell> patternCells = getPatternCells(row);
		List list = wrapper(data);
		int dataSize = list.size();
		for (int i = 0; i < generateLength; i++) {
			// 插入新行
			XWPFTableRow currentRow = table.insertNewTableRow(2 + i);
			// 复制行样式
			PoiStyleCopyUtils.copyRowStyle(row, currentRow);
			// 超过数据容量，则只填充空行
			if (i >= dataSize) {
				continue;
			}
			Object obj = list.get(i);
			// 为cell填充值
			for (int j = 0; j < patternCells.size(); j++) {
				PatternCell cell = patternCells.get(j);
				String value = cell.getName();
				// 如果是模式匹配，则通过EL计算值
				if (cell.isPattern()) {
					value = parser.parseExpression(value).getValue(obj, String.class);
				}
				XWPFTableCell currentCell = currentRow.getCell(j);
				currentCell.setText(value);
			}
		}
		// 删除第二行的模板
		table.removeRow(1);
	}

	private List<PatternCell> getPatternCells(XWPFTableRow patternRow) {
		return patternRow.getTableCells().stream().map(cell -> {
			String text = cell.getText();
			Matcher matcher = PATTERN.matcher(text);
			if (matcher.find()) {
				String group = matcher.group(1);
				return new PatternCell(true, group, cell);
			}
			return new PatternCell(true, text, cell);
		}).collect(Collectors.toList());
	}

	private List wrapper(Object data) {
		List list;
		if (data instanceof List) {
			list = (List) data;
		} else {
			list = Collections.singletonList(data);
		}
		return list;
	}

	@Override
	public String getName() {
		return NAME;
	}

}
