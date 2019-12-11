package com.deepoove.poi.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.exception.RenderException;
import com.deepoove.poi.template.ElementTemplate;
import com.deepoove.poi.template.run.RunTemplate;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.ByteArrayInputStream;
import java.util.Base64;

/**
 * 图片模式匹配RenderPolicy
 * image|height*width:var
 *
 * @author 奔波儿灞
 * @since 1.0
 */
public class ImagePatternRenderPolicy extends AbstractPatternRenderPolicy {

	public static final String NAME = "image";

	/**
	 * 单位cm
	 */
	private float width;

	/**
	 * 单位cm
	 */
	private float height;

	@Override
	public void render(ElementTemplate eleTemplate, Object data, XWPFTemplate template) {
		// 解析属性
		parseAttr();
		RunTemplate runTemplate = (RunTemplate) eleTemplate;
		XWPFRun run = runTemplate.getRun();
		run.setText("", 0);
		byte[] bytes = dataToBytes(data);
		// 猜测¬类型
		int suggestFileType = suggestFileType(bytes);
		try {
			run.addPicture(new ByteArrayInputStream(bytes), suggestFileType, "Generated", cm2emu(width), cm2emu(height));
		} catch (Exception e) {
			throw new RenderException("Image render error", e);
		}
	}

	private void parseAttr() {
		String attr = getAttr();
		int index = StringUtils.indexOf(attr, "*");
		if (index == -1) {
			throw new RenderException("Image render attr must be height*width");
		}
		String heightStr = attr.substring(0, index);
		try {
			height = Float.valueOf(heightStr);
		} catch (NumberFormatException e) {
			throw new RenderException("Image render attr height must be number: " + heightStr);
		}
		String widthStr = attr.substring(index + 1);
		try {
			width = Float.valueOf(widthStr);
		} catch (NumberFormatException e) {
			throw new RenderException("Image render attr width must be number: " + widthStr);
		}
	}

	private byte[] dataToBytes(Object data) {
		if (data instanceof String) {
			String image = (String) data;
			byte[] bytes;
			try {
				bytes = Base64.getDecoder().decode(image);
			} catch (IllegalArgumentException e) {
				throw new RenderException("Image render data must be image base64 encoded string");
			}
			return bytes;
		}
		throw new RenderException("Image render data must be image base64 encoded string");
	}

	private int suggestFileType(byte[] bytes) {
		byte[] head = new byte[28];
		if (bytes.length <= 28) {
			throw new RenderException("Image render data is base64 encoded string, but not a invalid [JPEG、PNG] image");
		}
		System.arraycopy(bytes, 0, head, 0, 28);
		String hex = bytes2hex(head);
		if (hex.startsWith(Type.JPEG.value)) {
			return Document.PICTURE_TYPE_JPEG;
		} else if (hex.startsWith(Type.PNG.value)) {
			return Document.PICTURE_TYPE_PNG;
		} else {
			throw new RenderException("Unsupported image type, only JPEG、PNG supported");
		}
	}

	private String bytes2hex(byte[] bytes) {
		StringBuilder builder = new StringBuilder();
		for (byte b : bytes) {
			int v = b & 0xFF;
			String hv = Integer.toHexString(v);
			if (hv.length() < 2) {
				builder.append(0);
			}
			builder.append(hv.toUpperCase());
		}
		return builder.toString();
	}

	private int cm2emu(float value) {
		return (int) (value * 360000);
	}

	@Override
	public String getName() {
		return NAME;
	}

	enum Type {

		/**
		 * jpg、jpeg
		 */
		JPEG("FFD8FF"),

		/**
		 * png
		 */
		PNG("89504E47"),

		;

		/**
		 * 文件魔数
		 */
		private String value;

		Type(String value) {
			this.value = value;
		}

	}

}
