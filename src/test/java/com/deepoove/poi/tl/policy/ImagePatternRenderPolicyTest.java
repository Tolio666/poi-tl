package com.deepoove.poi.tl.policy;

import com.deepoove.poi.XWPFTemplate;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.HashMap;

/**
 * 测试{@link com.deepoove.poi.policy.ImagePatternRenderPolicy}
 *
 * @author 奔波儿灞
 * @since 1.0
 */
public class ImagePatternRenderPolicyTest {

	private static final Logger LOG = LoggerFactory.getLogger(ImagePatternRenderPolicyTest.class);

	@Test
	public void run() throws IOException {
		Path logoPath = Paths.get("src/test/resources", "logo.png");
		byte[] bytes = Files.readAllBytes(logoPath);
		byte[] encode = Base64.getEncoder().encode(bytes);

		Path path = Paths.get("src/test/resources", "image_pattern.docx");
		XWPFTemplate template = XWPFTemplate.compile(path.toFile())
			// 数据
			.render(new HashMap<String, Object>() {{
				put("logo", new String(encode));
			}});
		// 输出
		Path outPath = Paths.get("src/test/resources", "image_pattern_out.docx");
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

}
