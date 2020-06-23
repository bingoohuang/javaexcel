package com.github.gobars.xlsx;

/**
 * JavaBean自定义校验接口.
 *
 * <p>用于JavaBean自定义校验
 */
public interface XlsxValidatable<T> {
  /**
   * 校验.
   *
   * @param lastError 基本校验消息(null表示基本校验没有错误)
   * @param message 当前校验的JavaBean或者Map
   * @return null 表示校验通过, 否则表示校验失败
   */
  String validate(String lastError, T message);
}
