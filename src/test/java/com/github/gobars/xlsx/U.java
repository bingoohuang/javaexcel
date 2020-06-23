package com.github.gobars.xlsx;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

@SuppressWarnings("unchecked")
public class U {
  public static <T> ArrayList<T> listOf(T... values) {
    return new ArrayList<>(Arrays.asList(values));
  }

  public static <T> Map<T, T> mapOf(T... values) {
    Map<T, T> m = new HashMap<>(values.length / 2 + 1);

    for (int i = 0; i < values.length; i += 2) {
      m.put(values[i], values[i + 1]);
    }

    return m;
  }
}
