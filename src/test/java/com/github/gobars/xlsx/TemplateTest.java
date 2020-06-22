package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.google.common.truth.Truth.assertThat;

public class TemplateTest {
  @Test
  public void test() {
    List<RowBean3> beans = new ArrayList<>();
    beans.add(new RowBean3().name("黄进兵").city("海淀区"));
    beans.add(new RowBean3().name("兵进黄").city("西城区"));

    String excel = "excels/test-template.xlsx";

    new Xlsx().template("template.xlsx", FileType.CLASSPATH).fromBeans(beans).write(excel);

    List<RowBean3> read = new Xlsx().read(excel).toBeans(RowBean3.class);

    assertThat(read).isEqualTo(beans);
  }

  @Test
  public void horizontal() {
    List<TitleBean> bs = new ArrayList<>();
    bs.add(new TitleBean().title("地区").sample("示例-海淀区").d1("西城").d2("东城").d3("南城"));
    bs.add(new TitleBean().title("血压").sample("示例-140/90").d1("135/90").d2("140/95").d3("133/85"));
    bs.add(new TitleBean().title("性别").sample("示例-女").d1("男").d2("女").d3("未知"));
    bs.add(new TitleBean().title("学校").sample("示例-蓝翔").d1("东大").d2("西大").d3("北大"));

    new Xlsx()
        .template("template-horizontal.xlsx", FileType.CLASSPATH)
        .fromBeans(bs, new FromOption().horizontal(true))
        .write("excels/test-horizontal.xlsx");

    List<HorizontalBean> read =
        new Xlsx().read("excels/test-horizontal.xlsx").toBeans(HorizontalBean.class);
    assertThat(read)
        .containsExactly(
            new HorizontalBean().area("示例-海淀区").blood("示例-140/90").gender("示例-女").school("示例-蓝翔"),
            new HorizontalBean().area("西城").blood("135/90").gender("男").school("东大"),
            new HorizontalBean().area("东城").blood("140/95").gender("女").school("西大"),
            new HorizontalBean().area("南城").blood("133/85").gender("未知").school("北大"));

    List<TitleInfo> titleInfos = new ArrayList<>();
    titleInfos.add(new TitleInfo().title("地区").mapKey("area"));
    titleInfos.add(new TitleInfo().title("性别").mapKey("gender"));
    titleInfos.add(new TitleInfo().title("学校").mapKey("school"));
    titleInfos.add(new TitleInfo().title("血压").mapKey("blood"));

    List<Map<String, String>> maps =
        new Xlsx().read("excels/test-horizontal.xlsx").toBeans(titleInfos);
    assertThat(maps)
        .containsExactly(
            mapOf("area", "示例-海淀区", "blood", "示例-140/90", "gender", "示例-女", "school", "示例-蓝翔"),
            mapOf("area", "西城", "blood", "135/90", "gender", "男", "school", "东大"),
            mapOf("area", "东城", "blood", "140/95", "gender", "女", "school", "西大"),
            mapOf("area", "南城", "blood", "133/85", "gender", "未知", "school", "北大"));

    new Xlsx()
        .template("template-horizontal.xlsx", FileType.CLASSPATH)
        .fromBeans(bs, new FromOption().horizontal(true))
        .protect("123456")
        .write("excels/test-horizontal-123456.xlsx");

    List<HBean> hbeans = new Xlsx().read("excels/test-horizontal.xlsx").toBeans(HBean.class);
    assertThat(hbeans)
        .containsExactly(
            new HBean().area("示例-海淀区").blood("示例-140/90").rownum(4),
            new HBean().area("西城").blood("135/90").rownum(5),
            new HBean().area("东城").blood("140/95").rownum(6),
            new HBean().area("南城").blood("133/85").rownum(7));

    List<IgnoreBean> ibeans =
        new Xlsx().read("excels/test-horizontal.xlsx").toBeans(IgnoreBean.class);
    assertThat(ibeans)
        .containsExactly(
            new IgnoreBean().area("西城"), new IgnoreBean().area("东城"), new IgnoreBean().area("南城"));
  }

  private Map<String, String> mapOf(String... values) {
    Map<String, String> m = new HashMap<>(values.length / 2 + 1);

    for (int i = 0; i < values.length; i += 2) {
      m.put(values[i], values[i + 1]);
    }

    return m;
  }

  @Data
  @Accessors(fluent = true)
  public static class HBean implements RownumAware {
    @XlsxCol(title = "地区")
    private String area;

    @XlsxCol(title = "血压")
    private String blood;

    private int rownum;

    @Override
    public void setRownum(int rownum) {
      this.rownum = rownum;
    }
  }

  @Data
  @Accessors(fluent = true)
  public static class IgnoreBean implements IgnoreAware {
    @XlsxCol(title = "地区")
    private String area;

    @Override
    public boolean shouldIgnored() {
      return area != null && area.startsWith("示例");
    }
  }

  @Data
  @Accessors(fluent = true)
  public static class HorizontalBean {
    @XlsxCol(title = "地区")
    private String area;

    @XlsxCol(title = "血压")
    private String blood;

    @XlsxCol(title = "性别")
    private String gender;

    @XlsxCol(title = "学校")
    private String school;
  }

  @Data
  @Accessors(fluent = true)
  public static class TitleBean {
    @XlsxCol(title = "数据1")
    private String d1;

    @XlsxCol(title = "数据2")
    private String d2;

    @XlsxCol(title = "数据3")
    private String d3;

    @XlsxCol(title = "标题")
    private String title;

    @XlsxCol(title = "示例")
    private String sample;
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean3 {
    @XlsxCol(title = "姓名")
    private String name;

    @XlsxCol(title = "城市")
    private String city;
  }
}
