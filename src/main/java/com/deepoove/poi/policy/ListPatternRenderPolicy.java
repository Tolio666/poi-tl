package com.deepoove.poi.policy;

import com.deepoove.poi.NiceXWPFDocument;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

import java.util.Collections;
import java.util.List;

/**
 * 列表模式匹配RenderPolicy
 * list|limit:var
 *
 * @author 奔波儿灞
 * @since 1.0
 */
public class ListPatternRenderPolicy extends AbstractPatternRenderPolicy {

	public static final String NAME = "list";

	private int generateLength;

	@Override
	public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
		parseAttr();
		List list = wrapper(data);

		NiceXWPFDocument document = template.getXWPFDocument();
		RunTemplate runTemplate = (RunTemplate) eleTemplate;
		XWPFRun run = runTemplate.getRun();
		Integer pos = runTemplate.getRunPos();
		XWPFParagraph paragraph = (XWPFParagraph) runTemplate.getRun().getParent();

		int dataSize = list.size();
		// 如果为-1，则根据实际的数据条数展示
		if (generateLength == -1) {
			generateLength = dataSize;
		}
		for (int i = 0; i < generateLength; i++) {
			// 超过数据容量，则只填充空行
			if (i >= dataSize) {
				XWPFParagraph newParagraph = insertNewParagraphAfter(document, paragraph);
				// 设置段落样式
				newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
				continue;
			}
			Object obj = list.get(i);
			XWPFParagraph newParagraph = insertNewParagraphAfter(document, paragraph);
			// 复制
			copyParagraph(paragraph, newParagraph);
			// 移除占位
			newParagraph.removeRun(pos);
			// 新值替换
			XWPFRun newRun = newParagraph.insertNewRun(pos);
			newRun.setText(String.valueOf(obj));
			// 设置样式
			newRun.getCTR().setRPr(run.getCTR().getRPr());
		}
		// 移除第一行
		document.removeBodyElement(document.getPosOfParagraph(paragraph));
	}

	private void parseAttr() {
		String attr = getAttr();
		try {
			generateLength = Integer.valueOf(attr);
		} catch (NumberFormatException e) {
			throw new RenderException("List render attr must be number: " + attr);
		}
	}

	private List wrapper(Object data) {
		if (data instanceof List) {
			return (List) data;
		}
		return Collections.singletonList(data);
	}

	private XWPFParagraph insertNewParagraphAfter(XWPFDocument document, XWPFParagraph paragraph) {
		XmlCursor cursor = paragraph.getCTP().newCursor();
		return document.insertNewParagraph(cursor);
	}

	private void copyParagraph(XWPFParagraph source, XWPFParagraph target) {
		for (XWPFRun sourceRun : source.getRuns()) {
			XWPFRun targetRun = target.createRun();
			copyRun(sourceRun, targetRun);
		}
		// 设置段落样式
		target.getCTP().setPPr(source.getCTP().getPPr());
	}

	private void copyRun(XWPFRun source, XWPFRun target) {
		// 设置文本
		target.setText(source.text());
		// 设置样式
		target.getCTR().setRPr(source.getCTR().getRPr());
	}

	@Override
	public String getName() {
		return NAME;
	}

}
