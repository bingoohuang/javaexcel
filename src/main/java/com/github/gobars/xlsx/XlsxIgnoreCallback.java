package com.github.gobars.xlsx;

public interface XlsxIgnoreCallback<T> {
  boolean shouldIgnore(T bean);
}
