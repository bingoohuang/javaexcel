# xlsx

[![Build Status](https://travis-ci.org/gobars/xlsx.svg?branch=master)](https://travis-ci.org/gobars/xlsx)
[![Quality Gate](https://sonarcloud.io/api/project_badges/measure?project=com.github.gobars%3Axlsx&metric=alert_status)](https://sonarcloud.io/dashboard/index/com.github.gobars%3Axlsx)
[![Coverage Status](https://coveralls.io/repos/github/gobars/xlsx/badge.svg?branch=master)](https://coveralls.io/github/gobars/xlsx?branch=master)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.gobars/xlsx/badge.svg?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/com.github.gobars/xlsx/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

binding between java beans and excel rows based on poi.

## Usage

### JavaBean读取

![image](https://user-images.githubusercontent.com/1940588/85396427-51265a80-b584-11ea-8ddf-c6bf39c5ed2b.png)

定义与如上Excel对应的JavaBean：

```java
@Data
@Accessors(fluent = true)
public static class Bean {
    @XlsxCol(title = "地区", ignoreRow = "示例-") private String area;
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
        .fromBeans(beans, new XlsxOptionFrom().horizontal(true))
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

### JavaBean校验

1. 在JavaBean上定义注解`@XlsxValid`
1. 字段上使用Javax校验注解，或者Hibernate校验
1. 实现`XlsxValidatable`接口自定义业务校验的其它校验

例如:

```java
@Data
@Accessors(fluent = true)
@XlsxValid(writeErrorToExcel = true)
public static class XBean implements XlsxValidatable<XBean> {
    @NotNull
    @Length(max = 3)
    @XlsxCol(title = "地区", ignoreRow = "示例-")
    private String area;

    @XlsxCol("血压")
    private String blood;
    
    @XlsxCol("性别")
    @Pattern(regexp = "男|女")
    private String gender;

    @Override
    public String validate(String error, XBean bean) {
      // 自定义其它校验，比如业务校验，状态校验、存在校验、组合校验等
      // 返回null表示校验通过，否则返回具体校验失败信息
      return null;
    }
}
```

```java
List<XBean> vBeans = xlsx.toBeans(XBean.class);
// 因为XBean上注解了writeErrorToExcel = true，所以可以写出带有错误提示的excel
xlsx.write("excels/test-validate.xlsx"); 
```

### 读出到Map的校验

由于Map类型，无法使用注解，只能手动定义回调函数，例如.

```java
Xlsx xlsx = new Xlsx().read("validate.xlsx", CLASSPATH);

XlsxValidatable<Map<String, String>> validatable =
    (lastError, map) -> {
      String gender = map.get("gender");
      if ("男".equals(gender) || "女".equals(gender)) {
        return null;
      }

      return "性别格式错误，必须为男或女";
    };

XlsxIgnoreable<Map<String, String>> ignoreable =
    bean -> Util.contains(bean.get("area"), "示例-");

XlsOptionTo optionTo =
    new XlsOptionTo()
        // 将错误标识在Excel行末
        .writeErrorToExcel(true)
        .ignoreable(ignoreable)
        // 校验回调
        .validatable(validatable);

List<XlsxTitle> titles =
    XlsxTitle.create(mapOf("地区", "area", "性别", "gender", "血压", "blood"));

List<Map<String, String>> maps = xlsx.toBeans(titles, optionTo);
xlsx.write("excels/test-validate-map.xlsx");
```

更详细参见[ValidationTest.readExcel2Maps](src/test/java/com/github/gobars/xlsx/ValidateTest.java).

错误提示的excel:

![image](https://user-images.githubusercontent.com/1940588/85397426-1b827100-b586-11ea-9c59-cfa7c140078b.png)
