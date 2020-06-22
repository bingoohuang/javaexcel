package com.github.gobars.xlsx;

public enum FileType {
  /**
   * 类路径中的文件.
   *
   * <p>主要用于定义excel模板文件等需要一起打包发布的文件.
   */
  CLASSPATH,
  /**
   * 普通文件系统文件.
   *
   * <p>主要用于测试，或者指定文件系统的操作
   */
  NORMAL
}
