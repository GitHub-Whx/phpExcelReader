<?php
require_once '../vendor/autoload.php';
require_once '../src/ImportUtils.php';
use whx\phpExcelReader\ImportUtils;

class Demo{


    public function foo(){
        try {
            $utils = new ImportUtils();
            return $utils->readFile('./template.xlsx')
                ->setRowIndex(2)
                ->setDateFormat('Y-m-d')
                ->setDateColumnName(['Date'])
                ->run();
        }catch (\Exception $e){
            print_r($e->getMessage());
            exit();
        }

    }

}
$obj = new Demo();
echo json_encode($obj->foo());
exit();