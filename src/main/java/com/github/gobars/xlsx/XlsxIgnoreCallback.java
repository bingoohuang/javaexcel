package com.github.gobars.xlsx;

/**
 * 是否忽略当前bean.
 *
 * @param <T> Beasn类型
 * @author bingoo
 */
public interface XlsxIgnoreCallback<T> {
  /**
   * 是否忽略当前Bean.
   *
   * @param bean 当前Bean
   * @return true 忽略
   */
  boolean shouldIgnore(T bean);
}
