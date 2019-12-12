# Poi-tl(Poi-template-language)

> 基于 poi-tl 的二次开发

## 说明

修改了源码，新增 `AbstractPatternRenderPolicy` 抽象类，用于实现 `name|attr:var` 语法

其中：

- name 为名称
- attr 为一些属性
- var 为数据变量参数

注意：

- 默认开启 Spring EL 表达式
- 最大程度保证与 [poi-tl](https://github.com/Sayi/poi-tl) 的兼容性，目前基于 v1.6.0 fork， v1.6.0 的功能均能正常使用。

## 扩展

### 表格

语法：`{{table|limit:var}}`

- table说明是表格
- limit 为数据填充的行数，数据不足补空
- var 为填充数据（JSON）的 key，可以是一个数组。

模板：

![extend-table](docs/assets/extend-table.jpg)

其中：

- 姓名的前面出现的`{{table|5:[users]}}`，代表了这是一个表格模板，`users`则说明 JSON 数据中存在一个 users 的 key 。
- 表格的第二行变量会根据传递的值动态替换，{name}、{age} 等模板，则说明 users 这个 key 中的 JSON 对象存在 name、age 这两个key。
- 由于数据只有2条，限制5条，因此补空行3条

测试代码：

```java
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
```

可以看到这里的 JSON 对象（Java中可以是一个hashmap）存在 users 这个 key，且存在 2 条数据。User 这个对象有两个属性 name、age ，模板在解析时，会自动取值。

输出：

![extend-table-out](docs/assets/extend-table-out.jpg)

总结：表格正常渲染，而且样式也正常保留，原来的数据也会保留下来，数据不足补空行。

### 图片

语法：`{{image|height*width:var}}`

- image说明是图片
- height*width代表图片的高度和宽度，单位为厘米
- var为填充数据（JSON）的 key，是一个图片字节通过base64加密的字符串

模板：

![extend-image](docs/assets/extend-image.jpg)

测试代码：

```java
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
```

输出：

![extend-image-out](docs/assets/extend-image-out.jpg)

总结：图片能正常根据高度宽度渲染出来

### 列表

语法：`list|limit:var`

- list说明是列表
- limit 为数据填充的行数，数据不足补空
- var 为填充数据（JSON）的 key，值可以是一个字符串或者一个字符串数组。

模板：

![extend-list](docs/assets/extend-list.jpg)

测试代码：

```java
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
```

输出：

![extend-list-out](docs/assets/extend-list-out.jpg)

总结：列表样式支持罗马字符、有序无序等，代码里面指定列表样式。



