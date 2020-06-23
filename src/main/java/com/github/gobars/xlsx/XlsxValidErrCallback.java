package com.github.gobars.xlsx;

/**
 * 校验错误回调.
 *
 * @param <T> JavaBean类型.
 */
public interface XlsxValidErrCallback<T> {
  void call(T t, String errMsg, int rownum);
}
