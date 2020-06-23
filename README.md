# xlsx

[![Build Status](https://travis-ci.org/gobars/xlsx.svg?branch=master)](https://travis-ci.org/gobars/xlsx)
[![Quality Gate](https://sonarcloud.io/api/project_badges/measure?project=com.github.gobars%3Axlsx&metric=alert_status)](https://sonarcloud.io/dashboard/index/com.github.gobars%3Axlsx)
[![Coverage Status](https://coveralls.io/repos/github/gobars/xlsx/badge.svg?branch=master)](https://coveralls.io/github/gobars/xlsx?branch=master)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.gobars/xlsx/badge.svg?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/com.github.gobars/xlsx/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

binding between java beans and excel rows based on poi.

## Usage

### JavaBean读取

定义JavaBean：

```java
@Data
@Accessors(fluent = true)
public static class Bean {
    @XlsxCol("地区") private String area;
    @XlsxCol("血压") private String blood;
    @XlsxCol("性别") private String gender;
    @XlsxCol("学校") private String school;
}
```

读取到JavaBean列表中：

```java
List<Bean> read = new Xlsx().read("excels/test-horizontal.xlsx").toBeans(Bean.class);
```

### Map列表读取

```java
List<TitleInfo> titleInfos = new ArrayList<>();
titleInfos.add(new TitleInfo().title("地区").mapKey("area"));
titleInfos.add(new TitleInfo().title("性别").mapKey("gender"));
titleInfos.add(new TitleInfo().title("学校").mapKey("school"));
titleInfos.add(new TitleInfo().title("血压").mapKey("blood"));

List<Map<String, String>> maps = new Xlsx().read("excels/test-horizontal.xlsx").toBeans(titleInfos);
// maps 值为:  mapOf("area" => "南城", "blood" => "133/85", "gender" => "未知", "school" => "北大"));
```

### 横向生成

![image](https://user-images.githubusercontent.com/1940588/85288833-f8de5280-b4c8-11ea-80e1-8526ea61e58b.png)

```java
@Test
public void horizontal() {
    List<TitleBean> beans = new ArrayList<>();
    beans.add(new TitleBean().title("地区").sample("示例-海淀区").d1("西城").d2("东城").d3("南城"));
    beans.add(new TitleBean().title("血压").sample("示例-140/90").d1("135/90").d2("140/95").d3("133/85"));
    beans.add(new TitleBean().title("性别").sample("示例-女").d1("男").d2("女").d3("未知"));
    beans.add(new TitleBean().title("学校").sample("示例-蓝翔").d1("东大").d2("西大").d3("北大"));

    new Xlsx()
        .read("template-horizontal.xlsx", FileType.CLASSPATH)
        .fromBeans(beans, new FromOption().horizontal(true))
        .write("excels/test-horizontal.xlsx");
}

@Data
@Accessors(fluent = true)
public class TitleBean {
    @XlsxCol("标题")  private String title;
    @XlsxCol("示例") private String sample;
    @XlsxCol("数据1") private String d1;
    @XlsxCol("数据2") private String d2;
    @XlsxCol("数据3") private String d3;
}
```
