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
public class TitleInfo {
  private String title;
  private String mapKey;

  public static List<TitleInfo> create(Map<String, String> map) {
    val l = new ArrayList<TitleInfo>(map.size());

    for (val i : map.entrySet()) {
      l.add(new TitleInfo().title(i.getKey()).mapKey(i.getValue()));
    }

    return l;
  }
}
