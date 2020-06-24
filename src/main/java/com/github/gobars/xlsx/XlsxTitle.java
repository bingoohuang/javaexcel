package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import lombok.val;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 标题关联信息.
 *
 * <p>用于动态Excel到Map中.
 *
 * @author bingoobjca
 */
@Data
@Accessors(fluent = true)
public class XlsxTitle {
  private String title;
  private String mapKey;

  public static List<XlsxTitle> create(Map<String, String> map) {
    val l = new ArrayList<XlsxTitle>(map.size());

    for (val i : map.entrySet()) {
      l.add(new XlsxTitle().title(i.getKey()).mapKey(i.getValue()));
    }

    return l;
  }
}
