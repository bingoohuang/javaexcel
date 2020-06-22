package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;

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
}
