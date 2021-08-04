## [phpExcelReader](https://github.com/GitHub-Whx/phpExcelReader)

## 概述
基于 `phpoffice/phpspreadsheet` 封装的、简单的 `Excel` 数据读取

## 支持的文件格式
```
xlsx
xls
csv
```

## 功能概述
- 支持链式调用，提高代码清晰度和开发效率
- 支持默认表格数据读取
- 支持指定表格索引读取数据 `getSheetDataByIndex()`
- 支持指定表格名称读取数据 `getSheetDataByName()`
- 支持指定多个表格索引读取数据 `getBatchSheetDataByIndex()`
- 支持指定多个表格名称读取数据 `getBatchSheetDataByName()`
- 支持指定读取表格的列名读取数据 `setSheetTargetColumnName([ 'Sheet1' => ['Tracking NO.', 'Signed Quantity', 'Date'])` 或者 'setSheetTargetColumnName(['Tracking NO.', 'Signed Quantity', 'Date']])'
- 支持指定从哪一行开始读取表格数据 `->setRowIndex(['Sheet1' => 2])` 或者 `->setRowIndex(2)`
- 默认返回以表格索引或表格名称为键的多维数组
- 默认返回的数据是以列名为 `key` ，单元格值作为 `value` 的数组
- 可以指定只返回带列索引的原始数据 `isRaw()->isRawDataWithCellRef()`
- 支持返回格式化后的数据和带列索引的原始数据
- 支持日期类列数据转换读取，并支持指定日期格式 `->setSheetDateColumnName(['Sheet1' => ['Date'],'工作表2' => ['生日']])->setDateFormat('Y-m-d H:i:s')`

## 安装方法
```
composer require whx/phpexcelreader
```
## 运行环境
- PHP 7.0.0 已上版本
- composer

## 示例

- 默认读取活跃工作表 `activeSheet` 格式化后的数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}

```
- 读取表格原始数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
       ->isRaw(true)
       ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}

```
- 读取表格原始数据-带列索引
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
       ->isRaw(true)
       ->isRawDataWithCellRef(true)
       ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}

```

- 读取指定表格索引的数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getSheetDataByIndex(1)
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}

```
- 读取指定表格名称的数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getSheetDataByName('工作表2')
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}

```

- 读取多个指定表格索引的数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getBatchSheetDataByIndex([1,3])
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}

```
- 读取多个指定表格索引的数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getBatchSheetDataByName(['工作表2',Sheet1])
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}
```
- 指定表格名称、从哪一行开始读取、读取哪一列
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getBatchSheetDataByName(['工作表3','工作表2',Sheet1])
        ->setSheetTargetColumnName([
            'Sheet1'=>['Tracking NO.','Signed Quantity']
        ])
        ->setRowIndex([
            'Sheet1'=>2
        ])
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}
```
- 读取数据中，包含格式化数据和原始数据
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getBatchSheetDataByName(['工作表3','工作表2',Sheet1])
        ->setSheetTargetColumnName([
            'Sheet1'=>['Tracking NO.','Signed Quantity']
        ])
        ->setRowIndex([
            'Sheet1'=>2
        ])
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}
```
- 指定日期列，指定返回的日期格式
```
<?php
use whx\phpExcelReader\ExcelReader;
…………
try {
    $excelReader = new ExcelReader();
    return $excelReader->readFile('./template.xlsx')
        ->getBatchSheetDataByName(['工作表3','工作表2','Sheet1'])
        ->setSheetTargetColumnName([
            'Sheet1'=>['Tracking NO.','Signed Quantity','Date']
        ])
        ->setRowIndex([
            'Sheet1'=>2
        ])
        ->setSheetDateColumnName([
            'Sheet1'=>['Date'],
            '工作表2'=>['生日'],
        ])
        ->setDateFormat('Y-m-d H:i:s')
        ->run();
}catch (\Exception $e){
    print_r($e->getMessage());
    exit();
}
```
## 问题反馈
请移步 [github](https://github.com/GitHub-Whx/phpExcelReader) 反馈问题 

## License
phpExcelReader is licensed under `MIT`