package com.github.gobars.xlsxtest;

import com.github.gobars.xlsx.*;
import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.List;
import java.util.Map;

import static com.github.gobars.xlsx.XlsxFileType.CLASSPATH;
import static com.github.gobars.xlsx.XlsxUtil.listOf;
import static com.github.gobars.xlsx.XlsxUtil.mapOf;
import static com.google.common.truth.Truth.assertThat;

public class XlsxTest {
  @Test
  public void rows() {
    List<RowBean> rows = listOf(new RowBean().name("黄进兵"));

    String excel = "excels/test-rowsbean.xlsx";
    new Xlsx().fromBeans(rows).write(excel);
    List<RowBean> read = new Xlsx().read(excel).toBeans(RowBean.class);

    assertThat(read).isEqualTo(rows);
  }

  @Test
  public void rows2() {
    List<RowBean2> rows = listOf(new RowBean2().name("黄进兵"));

    String excel = "excels/test-rowsbean2.xlsx";

    new Xlsx().style("style.xlsx", CLASSPATH).fromBeans(rows).write(excel);

    List<RowBean2> read = new Xlsx().read(excel).toBeans(RowBean2.class);

    assertThat(read).isEqualTo(rows);
  }

  @Test
  public void rows3() {
    List<RowBean3> rows =
        listOf(new RowBean3().name("黄进兵").city("海淀区"), new RowBean3().name("兵进黄").city("西城区"));

    String excel = "excels/test-fromBeans.xlsx";

    new Xlsx().style("style.xlsx", CLASSPATH).fromBeans(rows).write(excel);

    List<RowBean3> read = new Xlsx().read(excel).toBeans(RowBean3.class);

    assertThat(read).isEqualTo(rows);
  }

  @Test
  @SuppressWarnings("unchecked")
  public void fromMaps() {
    List<XlsxTitle> titleInfos =
        XlsxTitle.create(XlsxUtil.mapOf("地区", "area", "性别", "gender", "血压", "blood"));

    List<Map<String, String>> maps =
        listOf(
            mapOf("area", "示例-海淀区", "blood", "示例-140/90", "gender", "示例-女", "school", "示例-蓝翔"),
            mapOf("area", "西城", "blood", "135/90", "gender", "男", "school", "东大"),
            mapOf("area", "东城", "blood", "140/95", "gender", "女", "school", "西大"),
            mapOf("area", "南城", "blood", "133/85", "gender", "未知", "school", "北大"));

    new Xlsx()
        .style("style.xlsx", CLASSPATH)
        .fromBeans(titleInfos, maps, new XlsxOptionFrom().titleStyle("A1").dataStyle("A2"))
        .write("excels/test-fromMaps.xlsx");

    List<XlsxTitle> titleInfos2 =
        XlsxTitle.create(XlsxUtil.mapOf("城市", "area", "姓名", "blood", "性别", "gender"));
    new Xlsx()
        .read("template.xlsx", CLASSPATH)
        .fromBeans(titleInfos2, maps)
        .write("excels/test-fromMaps-template.xlsx");
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean {
    @XlsxCol("姓名")
    private String name;
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean2 {
    @XlsxCol(title = "姓名", titleStyle = "A1", dataStyle = "A2")
    private String name;
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean3 {
    @XlsxCol(title = "姓名", titleStyle = "A1", dataStyle = "A2")
    private String name;

    @XlsxCol(title = "城市")
    private String city;
  }
}
