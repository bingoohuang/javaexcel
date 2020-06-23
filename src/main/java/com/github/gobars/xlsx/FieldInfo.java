package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * JavaBean字段在Excel中定位信息.
 *
 * @author bingoobjca
 */
@Data
@Accessors(fluent = true)
class FieldInfo {
  private String title;
  private String ignoreRow;

  private CellStyle titleStyle;
  private CellStyle dataStyle;

  /**
   * 索引。
   *
   * <p>垂直模式时，为列索引
   *
   * <p>水平模式时，为行索引
   */
  private int index;
}
