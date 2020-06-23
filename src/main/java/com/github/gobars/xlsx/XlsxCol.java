package com.github.gobars.xlsx;

import java.lang.annotation.*;

/**
 * 用来标识JavaBean属性对应到excel文档的一列.
 *
 * @author bingoobjca
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface XlsxCol {
  /**
   * Excel中定位用的标题头行包含的关键字.
   *
   * @return 行标题
   */
  String title() default "";

  /**
   * 行标题的别名。用来简化标签使用。
   *
   * @return 行标题
   */
  String value() default "";

  /**
   * 包含指定值时，忽略行
   *
   * @return 包含值时，忽略行，例如: "示例-"
   */
  String ignoreRow() default "";

  /**
   * 标题样式在模板样式中参考单元格.
   *
   * @return 单元格引用，例如A1
   */
  String titleStyle() default "";

  /**
   * 数据样式在模板样式中参考单元格.
   *
   * @return 单元格引用，例如A2
   */
  String dataStyle() default "";
}
