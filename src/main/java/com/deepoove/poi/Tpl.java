package com.deepoove.poi;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * 模板
 *
 * @author 奔波儿灞
 * @since 1.6.0
 */
public final class Tpl {

	private static final Logger LOG = LoggerFactory.getLogger(Tpl.class);

	private XWPFTemplate template;

	private Tpl() {
	}

	public static <T> Tpl render(InputStream is, T model) {
		Tpl tpl = new Tpl();
		tpl.template = XWPFTemplate.compile(is).render(model);
		return tpl;
	}

	public void out(OutputStream os) throws IOException {
		try {
			template.write(os);
		} finally {
			try {
				template.close();
			} catch (IOException e) {
				LOG.error("close template failed", e);
			}
		}
	}

}
