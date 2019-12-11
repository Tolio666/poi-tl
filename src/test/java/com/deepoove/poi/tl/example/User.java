package com.deepoove.poi.tl.example;

import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * 用户模型
 *
 * @author 奔波儿灞
 * @since 1.0
 */
@Data
@AllArgsConstructor
public class User {
	private String name;
	private Integer age;
}
