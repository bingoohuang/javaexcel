package com.github.gobars.xlsx;

/**
 * 校验错误回调.
 *
 * @param <T> JavaBean类型.
 */
public interface XlsxValidErrCallback<T> {
  /**
   * 错误回调.
   *
   * @param t JavaBean或者Map
   * @param errMsg 错误消息(Hibernate校验等)
   * @param rownum 当前Excel行号(0-based)
   */
  void call(T t, String errMsg, int rownum);
}
