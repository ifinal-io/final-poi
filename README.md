![Final Poi](LOGO.png)

# Final Poi

[![GitHub Workflow Status](https://img.shields.io/github/workflow/status/final-projects/final-poi/ci)](https://github.com/final-projects/final-poi/actions?query=workflow%3Aci)
[![LICENSE](https://img.shields.io/github/license/final-projects/final-poi)](http://www.apache.org/licenses/LICENSE-2.0.html)
![Maven Central](https://img.shields.io/maven-central/v/org.ifinal.finalframework/final-poi?label=maven)
![RELEASE](https://img.shields.io/nexus/r/org.ifinal.finalframework/final-poi?label=realease&server=https%3A%2F%2Foss.sonatype.org%2F)
![SNAPSHOT](https://img.shields.io/nexus/s/org.ifinal.finalframework/final-poi?label=snapshot&server=https%3A%2F%2Foss.sonatype.org%2F)

Simple excel generate tool

## Feature

* 动态列
* 表达式，支持Excel原生函数
* 多级表头或表尾
* 自定义样式，如背景颜色、字体等

## Usage

### Import Dependency

```xml

<dependency>
    <groupId>org.ifinal.finalframework</groupId>
    <artifactId>final-poi</artifactId>
    <version>${latest.version}</version>
</dependency>
```

### Configure Excel

```json
{
  "version": "XLSX",
  "styles": [
    {
      "title": "标题",
      "locked": true,
      "horizontalAlignment": "CENTER",
      "verticalAlignment": "CENTER",
      "foregroundColor": "RED",
      "font": {
        "name": "楷体",
        "size": 24,
        "color": "#00FF00"
      }
    }
  ],
  "sheets": [
    {
      "name": "SHEET NAME",
      "defaultRowHeight": 30,
      "defaultColumnWidth": -1,
      "headers": [
        {
          "cells": [
            {
              "style": "标题",
              "value": "姓名"
            },
            {
              "style": "标题",
              "value": "生日"
            },
            {
              "style": "标题",
              "value": "年龄"
            },
            {
              "style": "标题",
              "value": "生日/2"
            }
          ]
        }
      ],
      "body": {
        "cells": [
          {
            "columnWidth": 30,
            "value": "#{name}"
          },
          {
            "value": "#{age}"
          },
          {
            "columnWidth": 30,
            "value": "#{new java.text.SimpleDateFormat('yyyy-MM-dd hh:mm:ss').format(birthday)}"
          },
          {
            "value": "=B#{#cell.rowIndex + 1}/2"
          }
        ]
      },
      "footers": [
        {
          "cells": [
            {
              "value": "平均年龄："
            },
            {
              "value": "=AVERAGE(B2:B#{#cell.rowIndex})"
            }
          ]
        }
      ]
    }
  ]
}
```

### Write

```java
package org.ifinal.finalframework.poi;

class ExcelsTest {

    private List<Person> persons() {
        return Arrays.asList(
            new Person("xiaoMing", 12, new Date()),
            new Person("xiaoHong", 18, new Date())
        );
    }

    @Test
    void testJson() throws IOException {

        final String json = "...";

        Excel excel = objectMapper.readValue(json, Excel.class);

        Excels.newWriter(excel).append(persons())
            .append(persons())
            .write("test3.xlsx");

    }

}
```

![excel generate](static/images/excel.generate.png)

## Maven Repositories

因Maven发布存在延迟，若从Maven中心下载不到，可添加以下配置从OSS仓库下载：

```xml

<repositories>
    <repository>
        <id>oss</id>
        <name>oss-releases</name>
        <url>https://oss.sonatype.org/content/repositories/releases/</url>
        <releases>
            <enabled>true</enabled>
        </releases>
        <snapshots>
            <enabled>false</enabled>
        </snapshots>
    </repository>
    <repository>
        <id>oss-snapshots</id>
        <name>oss-snapshots</name>
        <url>https://oss.sonatype.org/content/repositories/snapshots/</url>
        <releases>
            <enabled>false</enabled>
        </releases>
        <snapshots>
            <enabled>true</enabled>
        </snapshots>
    </repository>
</repositories>
```