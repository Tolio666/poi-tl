package com.deepoove.poi.policy;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 抽象的模式匹配RenderPolicy，用于实现类似的需求name|attr:var
 *
 * @author 奔波儿灞
 * @since 1.6.0
 */
public abstract class AbstractPatternRenderPolicy implements RenderPolicy {

	private static final Pattern PATTERN = Pattern.compile("(.+)\\|(.+):(.+)");

	private String attr;
	private String var;

	public boolean support(String tagName) {
		Pattern pattern = Pattern.compile(getName() + "\\|(.+):(.+)");
		Matcher matcher = pattern.matcher(tagName);
		if (matcher.find()) {
			attr = matcher.group(1);
			var = matcher.group(2);
			return true;
		}
		return false;
	}

	public abstract String getName();

	public String getAttr() {
		return attr;
	}

	public String getVar() {
		return var;
	}

	public static boolean isPattern(String tagName) {
		Matcher matcher = PATTERN.matcher(tagName);
		return matcher.find();
	}

}
