package com.github.gobars.xlsx;

import lombok.Data;
import lombok.experimental.Accessors;
import org.junit.Test;

import java.util.List;
import java.util.Map;

import static com.github.gobars.xlsx.FileType.CLASSPATH;
import static com.github.gobars.xlsx.Util.listOf;
import static com.github.gobars.xlsx.Util.mapOf;
import static com.google.common.truth.Truth.assertThat;

public class TemplateTest {
  @Test
  public void test() {
    List<RowBean3> beans =
        listOf(new RowBean3().name("黄进兵").city("海淀区"), new RowBean3().name("兵进黄").city("西城区"));

    String excel = "excels/test-template.xlsx";

    new Xlsx().autoClose(true).read("template.xlsx", CLASSPATH).fromBeans(beans).write(excel);

    List<RowBean3> read = new Xlsx().read(excel).toBeans(RowBean3.class);

    assertThat(read).isEqualTo(beans);
  }

  @Test
  public void horizontal() {
    List<TitleBean> bs =
        listOf(
            new TitleBean().title("地区").sample("示例-海淀区").d1("西城").d2("东城").d3("南城"),
            new TitleBean().title("血压").sample("示例-140/90").d1("135/90").d2("140/95").d3("133/85"),
            new TitleBean().title("性别").sample("示例-女").d1("男").d2("女").d3("未知"),
            new TitleBean().title("学校").sample("示例-蓝翔").d1("东大").d2("西大").d3("北大"));

    new Xlsx()
        .read("template-horizontal.xlsx", CLASSPATH)
        .fromBeans(bs, new OptionFrom().horizontal(true))
        .write("excels/test-horizontal.xlsx");

    List<HorizontalBean> read =
        new Xlsx().read("excels/test-horizontal.xlsx").toBeans(HorizontalBean.class);
    assertThat(read)
        .containsExactly(
            new HorizontalBean().area("示例-海淀区").blood("示例-140/90").gender("示例-女").school("示例-蓝翔"),
            new HorizontalBean().area("西城").blood("135/90").gender("男").school("东大"),
            new HorizontalBean().area("东城").blood("140/95").gender("女").school("西大"),
            new HorizontalBean().area("南城").blood("133/85").gender("未知").school("北大"));

    List<TitleInfo> titleInfos =
        listOf(
            new TitleInfo().title("地区").mapKey("area"),
            new TitleInfo().title("性别").mapKey("gender"),
            new TitleInfo().title("学校").mapKey("school"),
            new TitleInfo().title("血压").mapKey("blood"));

    List<Map<String, String>> maps =
        new Xlsx().read("excels/test-horizontal.xlsx").toBeans(titleInfos);
    assertThat(maps)
        .containsExactly(
            mapOf("area", "示例-海淀区", "blood", "示例-140/90", "gender", "示例-女", "school", "示例-蓝翔"),
            mapOf("area", "西城", "blood", "135/90", "gender", "男", "school", "东大"),
            mapOf("area", "东城", "blood", "140/95", "gender", "女", "school", "西大"),
            mapOf("area", "南城", "blood", "133/85", "gender", "未知", "school", "北大"));

    new Xlsx()
        .read("template-horizontal.xlsx", CLASSPATH)
        .fromBeans(bs, new OptionFrom().horizontal(true))
        .protect("123456")
        .write("excels/test-horizontal-123456.xlsx");

    List<IgnoreBean> ibeans =
        new Xlsx().read("excels/test-horizontal.xlsx").toBeans(IgnoreBean.class);
    assertThat(ibeans)
        .containsExactly(
            new IgnoreBean().area("西城"), new IgnoreBean().area("东城"), new IgnoreBean().area("南城"));
  }

  @Data
  @Accessors(fluent = true)
  public static class IgnoreBean {
    @XlsxCol(title = "地区", ignoreRow = "示例-")
    private String area;
  }

  @Data
  @Accessors(fluent = true)
  public static class HorizontalBean {
    @XlsxCol("地区")
    private String area;

    @XlsxCol("血压")
    private String blood;

    @XlsxCol("性别")
    private String gender;

    @XlsxCol("学校")
    private String school;
  }

  @Data
  @Accessors(fluent = true)
  public static class TitleBean {
    @XlsxCol("数据1")
    private String d1;

    @XlsxCol("数据2")
    private String d2;

    @XlsxCol("数据3")
    private String d3;

    @XlsxCol("标题")
    private String title;

    @XlsxCol("示例")
    private String sample;
  }

  @Data
  @Accessors(fluent = true)
  public static class RowBean3 {
    @XlsxCol("姓名")
    private String name;

    @XlsxCol("城市")
    private String city;
  }
}
