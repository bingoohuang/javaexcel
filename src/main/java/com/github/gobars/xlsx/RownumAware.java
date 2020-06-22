package com.github.gobars.xlsx;

/**
 * 感知当前Bean对应的Excel行.
 *
 * <p>行号是0-based.
 *
 * @author bingoo
 */
public interface RownumAware {
  /**
   * 设置行号.
   *
   * @param rownum 行号
   */
  void setRownum(int rownum);
}
