package com.deepoove.poi.tl.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.tl.example.User;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.standard.SpelExpressionParser;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

/**
 * 测试{@link com.deepoove.poi.policy.TablePatternRenderPolicy}
 *
 * @author 奔波儿灞
 * @since 1.0
 */
public class TablePatternRenderPolicyTest {

	private static final Logger LOG = LoggerFactory.getLogger(TablePatternRenderPolicyTest.class);

	@Test
	public void run() {
		Path path = Paths.get("src/test/resources", "table_pattern.docx");
		XWPFTemplate template = XWPFTemplate.compile(path.toFile())
			// 数据
			.render(new HashMap<String, Object>() {{
				put("users", Arrays.asList(new User("张三", 1), new User("李四", 2)));
			}});
		// 输出
		Path outPath = Paths.get("src/test/resources", "table_pattern_out.docx");
		try (OutputStream os = new BufferedOutputStream(new FileOutputStream(outPath.toFile()))) {
			template.write(os);
		} catch (IOException e) {
			LOG.error("render tpl error", e);
		} finally {
			try {
				template.close();
			} catch (IOException e) {
				LOG.error("close template error", e);
			}
		}
	}

	@Test
	public void el() {
		Map<String, Object> map = new HashMap<>();
		map.put("users", Arrays.asList(new User("张三", 1), new User("李四", 2)));
		ExpressionParser parser = new SpelExpressionParser();
		Object users = parser.parseExpression("[users]").getValue(map);
	}

}
