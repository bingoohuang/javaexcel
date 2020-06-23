package com.github.gobars.xlsx;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class U {
  public static <T> ArrayList<T> listOf(T... values) {
    ArrayList<T> l = new ArrayList<>(values.length);

    for (T v : values) {
      l.add(v);
    }

    return l;
  }

  public static Map<String, String> mapOf(String... values) {
    Map<String, String> m = new HashMap<>(values.length / 2 + 1);

    for (int i = 0; i < values.length; i += 2) {
      m.put(values[i], values[i + 1]);
    }

    return m;
  }
}
