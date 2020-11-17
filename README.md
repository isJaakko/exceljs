# 使用 exceljs 处理 Excel

## run
```
yarn build
```

## 背景
昨天朋友找我帮忙写一个处理 excel 的脚本，需求是：将表格中的正确答案标记为高亮状态。

初始表格长这样：
![origin](https://raw.githubusercontent.com/isJaakko/exceljs/main/static/screenshot_origin.png)

希望的结果：

![expect](https://raw.githubusercontent.com/isJaakko/exceljs/main/static/screenshot_expect.png)


## 整理需求
1. 熟悉表格内容

通过观察表格可以得知，每行的组成内容依次是：试题类型、试题内容、选项、正确答案

2. 找到正确答案

根据前置的信息，可以知道每行的第 `7` 列为正确答案

3. 高亮

已知信息全部拿到，接下来就是找到正确答案对应的单元格，给其设置样式即可。上图中的正确答案对应的单元格有：`C2`、`C3`、`F4`、`D5`。

## 开发过程

经过一番~~调研~~ Google，找到了 [exceljs](https://github.com/exceljs/exceljs/blob/master/README_zh.md) 这个库，大概看了下文档，发现符合我的需求，遂选用。

### 导入 excel
官方推荐使用[异步方式](https://github.com/exceljs/exceljs/blob/master/README_zh.md#%E5%8F%AF%E8%AF%BB%E6%B5%81)读取：
```
const filename = path.resolve(__dirname, "./file.xlsx");
const options = {
  sharedStrings: 'emit',
  hyperlinks: 'emit',
  worksheets: 'emit',
};

const workbookReader = new ExcelJS.stream.xlsx.WorkbookReader(filename, options);
workbookReader.read();

workbookReader.on('worksheet', worksheet => {
  worksheet.on('row', row => {
  });
});

workbookReader.on('shared-strings', sharedString => {
  // ...
});

workbookReader.on('hyperlinks', hyperlinksReader => {
  // ...
});

workbookReader.on('end', () => {
  // ...
});
workbookReader.on('error', (err) => {
  // ...
});
```

当然你也可以使用同步方式
```
const filename = path.resolve(__dirname, "./file.xlsx");
const options = {
    sharedStrings: "emit",
    hyperlinks: "emit",
    worksheets: "emit",
};

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(filename);

workbook.eachSheet(function (worksheet, sheetId) {
    worksheet.eachRow(function (row, rowNumber) {
        // ...
    };
}
```

### 找到所需单元格
对单元格的遍历中，根据标记正确答案的单元格，可以找到对应的正确答案所在单元格的行列。
```
workbook.eachSheet(function (worksheet, sheetId) {
    worksheet.eachRow(function (row, rowNumber) {
        // 正确答案在第 8 行
        const answers = row.getCell(8);
        // 匹配包含 ABCD 字符
        const pattern = /[A-D]/g;
        // 正确答案 ABCD 对应的列标号
        const answerMap = {
            A: "D",
            B: "E",
            C: "F",
            D: "G",
        };
        
        if (pattern.test(answers)) {
            // 存在多选情况
            answers
                .toString()
                .split("")
                .forEach(answer => {
                    const col = answerMap[answer];
                    // 找到正确答案对应的「列行号」，格式形如「A2」
                    const index = `${col}${rowNumber}`;
                }
        }
    }
}
```


### 修改单元格

这里碰到个坑，按照官方文档的写法，获取单个单元格对其设置样式会将所有单元格都修改。
```
worksheet.getCell('A1').fill = {
  type: 'pattern',
  pattern:'darkVertical',
  fgColor:{argb:'FFFF0000'}
};
```

折腾了好久，找到一个 [issues](https://github.com/exceljs/exceljs/issues/791)，改成了如下的写法，问题解决。
```
worksheet.getCell('A1').style = {
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF00" },
    bgColor: { argb: "FF00FF00" },
  },
};
```

本例中我修改了单元格填充，官方还支持许多其它操作，[文档](https://github.com/exceljs/exceljs/blob/master/README_zh.md)写得很清楚，这里就不一一列举了。

### 保存 excel
对 `excel` 操作完成后要写入才能对操作进行保存。
```
workbook.xlsx.writeFile(filename);
```

## 链接
- 本文源码：https://github.com/isJaakko/exceljs
- 官方文档：https://github.com/exceljs/exceljs
- 官方文档（中文）：https://github.com/exceljs/exceljs/blob/master/README_zh.md