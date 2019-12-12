package com.deepoove.poi.tl.policy;

import com.deepoove.poi.Tpl;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * 测试{@link com.deepoove.poi.policy.ListPatternRenderPolicy}
 *
 * @author 奔波儿灞
 * @since 1.0
 */
public class ListPatternRenderPolicyTest {

	private static final Logger LOG = LoggerFactory.getLogger(ListPatternRenderPolicyTest.class);

	@Test
	public void run() {
		Path inPath = Paths.get("src/test/resources", "list_pattern.docx");
		Path outPath = Paths.get("src/test/resources", "list_pattern_out.docx");
		Map<String, Object> model = new HashMap<String, Object>() {{
			put("items", Arrays.asList("张三", "李四", "王五"));
		}};
		try (InputStream is = Files.newInputStream(inPath); OutputStream os = Files.newOutputStream(outPath)) {
			Tpl.render(is, model).out(os);
		} catch (IOException e) {
			LOG.info("render tpl failed", e);
		}
	}

}
