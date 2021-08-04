<?php
require_once '../vendor/autoload.php';
require_once '../src/ExcelReader.php';

use whx\phpExcelReader\ExcelReader;

class Demo
{
    public function foo()
    {
        try {
            $excelReader = new ExcelReader();
            return $excelReader->readFile('./template.xlsx')
                ->getBatchSheetDataByName(['工作表3', '工作表2', 'Sheet1'])
                ->setSheetTargetColumnName([
                    'Sheet1' => ['Tracking NO.', 'Signed Quantity', 'Date']
                ])
                ->setRowIndex([
                    'Sheet1' => 2
                ])
                ->setSheetDateColumnName([
                    'Sheet1' => ['Date'],
                    '工作表2' => ['生日'],
                ])
                ->setDateFormat('Y-m-d H:i:s')
                ->run();
        } catch (\Exception $e) {
            print_r($e->getMessage());
            exit();
        }

    }
}

$obj = new Demo();
echo json_encode($obj->foo());
exit();