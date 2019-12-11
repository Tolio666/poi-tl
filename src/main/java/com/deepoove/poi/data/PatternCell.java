package com.deepoove.poi.data;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

/**
 * cell pattern data
 *
 * @author 奔波儿灞
 * @since 1.6.0
 */
@Data
@AllArgsConstructor
public class PatternCell {

	private boolean pattern;

	private String name;

	private XWPFTableCell cell;

}
