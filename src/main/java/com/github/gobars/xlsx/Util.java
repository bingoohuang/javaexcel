package com.github.gobars.xlsx;

import lombok.experimental.UtilityClass;
import lombok.extern.slf4j.Slf4j;

import java.io.Closeable;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

@Slf4j
@UtilityClass
class Util {
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

  public String getTitle(XlsxCol xlsxCol) {
    if (xlsxCol == null) {
      return "";
    }

    if (Util.isNotEmpty(xlsxCol.title())) {
      return xlsxCol.title();
    }

    return xlsxCol.value();
  }

  public boolean contains(String s, String sub) {
    return s != null && s.contains(sub);
  }

  public <T> ArrayList<T> listOf(T... values) {
    return new ArrayList<>(Arrays.asList(values));
  }

  public <T> HashMap<T, T> mapOf(T... values) {
    HashMap<T, T> m = new HashMap<>(values.length / 2 + 1);

    for (int i = 0; i < values.length; i += 2) {
      m.put(values[i], values[i + 1]);
    }

    return m;
  }
}
