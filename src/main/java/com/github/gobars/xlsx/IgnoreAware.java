package com.github.gobars.xlsx;

/**
 * 感知当前Bean是否需要被忽略.
 *
 * @author bingoo
 */
public interface IgnoreAware {
  /**
   * 是否忽略当前Bean.
   *
   * @return true 忽略
   */
  boolean shouldIgnored();
}
