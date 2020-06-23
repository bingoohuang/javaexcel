package com.github.gobars.xlsx;

import lombok.experimental.UtilityClass;

import java.io.Closeable;
import java.io.IOException;

@UtilityClass
public class Util {
  public boolean isEmpty(String s) {
    return s == null || s.length() == 0;
  }

  public boolean isNotEmpty(String s) {
    return s != null && s.length() > 0;
  }

  public void closeQuietly(Closeable closeable) {
    if (closeable != null) {
      try {
        closeable.close();
      } catch (IOException e) {
        // ignore
      }
    }
  }
}
