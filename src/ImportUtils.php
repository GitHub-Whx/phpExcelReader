<?php

/**
 * @Author: Whx
 * @Date  : 2021-07-22 18:46:50
 * @Last  Modified by:   Whx
 * @Last  Modified time: 2021-07-22 18:46:50
 */

namespace Whx\phpExcelReader;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

/**
 * Class ImportUtils
 */
/**
 * Class ImportUtils
 */
class ImportUtils
{

    /**
     * @var 输入的文件
     */
    private $_input_file;

    /**
     * @var array 支持的文件类型
     */
    private $_suport_suffix = ['csv', 'xls', 'xlsx'];

    /**
     * @var 获取到的文件实例
     */
    private $_spreadsheet;

    /**
     * @var 工作表 Sheet 总数
     */
    private $_sheet_count = 1;

    /**
     * @var array 工作表名称数组
     */
    private $_sheet_names = [];

    /**
     * @var array 获取到的工作表数据
     */
    private $_sheet_collection = [];

    /**
     * @var bool 是否获取多个 sheet
     */
    private $_is_multi_sheet = false;

    /**
     * @var int 默认的活动工作表序号
     */
    private $_default_active_sheet_index = 0;

    /**
     * @var bool 是否获取源数据
     */
    private $_is_raw = false;

    /**
     * @var bool 返回元数据时是否带上单元格信息
     */
    private $_is_raw_data_with_cell_ref = false;

    /**
     * @var bool 返回数据中是否带上源数据
     */
    private $_is_with_raw_data = false;

    /**
     * @var 当前结果集的键值
     */
    private $_curr_collection_key;

    /**
     * @var array 需要读取 sheet 的表头
     */
    private $_active_sheet_header = [];

    /**
     * @var array sheet 表从哪一行开始读取
     */
    private $_row_index = [];

    /**
     * @var int 当前 sheet 表从哪一行读取【没设置则认为第 1 行】
     */
    private $_curr_row_index = 1;

    /**
     * @var array sheet 表要读取的列名称【没设置则认为所有列】
     */
    private $_sheet_target_column_name = [];

    /**
     * @var array 当前 sheet 表要读取的列名称【没设置则认为所有列】
     */
    private $_curr_sheet_target_column_name = [];

    /**
     * @var array 当前 sheet 表要读取的列索引
     */
    private $_curr_sheet_target_column_index = [];

    /**
     * @var string 日期格式
     */
    private $_date_format = 'Y-m-d H:i:s';

    /**
     * @var array 日期类列名数组
     */
    private $_date_column_name = [];

    /**
     * @var array 结果
     */
    private $_response = [];


    /**
     * @param $file_name
     * Desc: 读取文件，并获取基本工作表的基本信息
     * User: Whx
     * Date: 2021-07-26 12:03
     */
    public function readFile($file_name)
    {

        $this->inputFileCheck($file_name);

        $this->_spreadsheet = IOFactory::load($this->_input_file);

        $this->_sheet_count = $this->_spreadsheet->getSheetCount();

        $this->_sheet_names = $this->_spreadsheet->getSheetNames();

        return $this;
    }


    /**
     * @param int $sheet_index
     * Desc: 通过工作表序号获取单个工作表的数据
     * User: Whx
     * Date: 2021-07-26 12:06
     */
    public function getSheetDataByIndex($sheet_index = 0)
    {
        // 是否为数值
        if (!is_numeric($sheet_index) || strpos($sheet_index, ".")) {

            throw new \Exception('sheet index must be an integer.');

        }

        // 序号是否有效校验
        if ($sheet_index >= ($this->_sheet_count)) {

            throw new \Exception('Your requested sheet index: ' . $sheet_index . ' is out of bounds. The actual bounds of sheets is ' . $this->_sheet_count);

        }

        try {

            $this->_sheet_collection[$sheet_index] = $this->_spreadsheet->getSheet($sheet_index);

            $this->_is_multi_sheet = false;

        } catch (\Exception $e) {

            throw new \Exception($e->getMessage());

        }


        return $this;
    }


    /**
     * @param $sheet_name
     * Desc: 通过工作表名称获取单个工作表的数据
     * User: Whx
     * Date: 2021-07-23 11:35
     */
    public function getSheetDataByName($sheet_name = 'Sheet1')
    {

        // 名称是否有效校验
        if (!in_array($sheet_name, $this->_sheet_names)) {

            throw new \Exception('Your requested sheet name: ' . $sheet_name . ' is out of bounds. The actual names of sheets is ' . json_encode($this->_sheet_name, JSON_UNESCAPED_UNICODE));

        }

        try {

            $this->_sheet_collection[$sheet_name] = $this->_spreadsheet->getSheetByName($sheet_name);

            $this->_is_multi_sheet = false;

        } catch (\Exception $e) {

            throw new \Exception($e->getMessage());

        }


        return $this;
    }


    /**
     * @param array $sheet_indexs
     * Desc: 通过指定工作表序号获取多个 sheet 数据
     * User: Whx
     * Date: 2021-07-23 17:30
     */
    public function getBatchSheetDataByIndex($sheet_indexs = [])
    {

        $allSheetData = $this->_spreadsheet->getAllSheets();

        foreach ($sheet_indexs as $sheet_index) {

            $this->_sheet_collection[$sheet_index] = $allSheetData[$sheet_index];

        }

        $this->_is_multi_sheet = true;

        return $this;
    }


    /**
     * @param array $sheet_names
     * Desc: 通过指定工作表名称获取多个 sheet 数据
     * User: Whx
     * Date: 2021-07-23 17:30
     */
    public function getBatchSheetDataByName($sheet_names = [])
    {

        $allSheetData = $this->_spreadsheet->getAllSheets();

        $worksheetCount = count($allSheetData);

        for ($i = 0; $i < $worksheetCount; $i++) {

            foreach ($sheet_names as $sheet_name) {

                if ($allSheetData[$i]->getTitle() === trim($sheet_name, "'")) {

                    $this->_sheet_collection[$sheet_name] = $allSheetData[$i];

                }
            }
        }

        $this->_is_multi_sheet = true;

        return $this;
    }


    /**
     * @return \app\cn\hofan\utils\ImportUtils
     * Desc: 获取所有工作表数据-默认以 sheet 的名称作为键值
     * User: Whx
     * Date: 2021-07-23 17:29
     */
    public function getAllSheetData()
    {

        $allSheetData = $this->_spreadsheet->getAllSheets();  // 数组

        $worksheetCount = count($allSheetData);

        for ($i = 0; $i < $worksheetCount; $i++) {

            foreach ($this->_sheet_name as $sheet_name) {

                if ($allSheetData[$i]->getTitle() === trim($sheet_name, "'")) {

                    $this->_sheet_collection[$sheet_name] = $allSheetData[$i];

                }
            }
        }

        $this->_is_multi_sheet = true;

        return $this;
    }

    /**
     * @param $is_raw
     * Desc: Description
     * User: Whx
     * Date: 2021-07-26 13:50
     */
    public function isRaw($is_raw)
    {
        $this->_is_raw = $is_raw;

        return $this;
    }

    /**
     * @param $_is_raw_data_with_cell_ref
     * Desc: Description
     * User: Whx
     * Date: 2021-07-26 13:52
     */
    public function isRawDataWithCellRef($_is_raw_data_with_cell_ref)
    {
        $this->_is_raw_data_with_cell_ref = $_is_raw_data_with_cell_ref;

        return $this;
    }

    /**
     * @param $is_with_raw_data
     * Desc: Description
     * User: Whx
     * Date: 2021-07-26 20:16
     */
    public function isWithRawData($is_with_raw_data)
    {
        $this->_is_with_raw_data = $is_with_raw_data;

        return $this;

    }

    /**
     * @param array $row_index
     * @return array
     */
    public function setRowIndex($row_index = [])
    {
        $this->_row_index = $row_index;

        return $this;
    }

    /**
     * @param array $sheet_target_column_name 【key=>value 形式】
     * Desc: 设置 sheet 要读取的列
     * User: Whx
     * Date: 2021-07-26 20:19
     */
    public function setSheetTargetColumnName($sheet_target_column_name = [])
    {

        $this->_sheet_target_column_name = $sheet_target_column_name;

        return $this;

    }

    /**
     * 设置日期格式
     * FuncName: setDateFormat
     * Author: Whx
     * DateTime: 2021-08-03 11:25:48
     * @param $format
     * @return $this
     */
    public function setDateFormat($format){

        $this->_date_format = $format;

        return $this;

    }

    /**
     * 设置日期类列名
     * FuncName: setDateColumnName
     * Author: Whx
     * DateTime: 2021-08-03 10:57:19
     * @param $date_column_name
     * @return $this
     */
    public function setDateColumnName($date_column_name){

        $this->_date_column_name = $date_column_name;

        return $this;
    }

    public function run()
    {

        if (!$this->_sheet_collection) {

            $this->getDefaultSheetData();

        }
        $this->_response = $this->sheetCollectionHandle();

        return $this->_response;
    }

    /**
     * Desc: Description
     * User: Whx
     * Date: 2021-07-26 13:44
     */
    private function sheetCollectionHandle()
    {

        $ret = [];
        foreach ($this->_sheet_collection as $key => $collection) {

            $this->_curr_collection_key = $key;

            if ($this->_is_raw) {

                $ret[$key] = $this->rawCollectionHandle($collection);

            } else {

                $ret[$key] = $this->collectionHandle($collection);

            }

        }

        return $ret;

    }

    /**
     * @param $collection
     * @return mixed
     * Desc: 源数据处理
     * User: Whx
     * Date: 2021-07-26 13:56
     */
    private function rawCollectionHandle($collection)
    {

        if ($this->_is_raw_data_with_cell_ref) {

            return $collection->toArray(null, true, true, true);

        }

        return $collection->toArray();

    }


    /**
     * @param $collection
     * @param $key_of_collection
     * Desc: Description
     * User: Whx
     * Date: 2021-07-26 20:26
     */
    private function collectionHandle($collection)
    {

        $ret = [];

        // 结果集转数组【带列名索引】
        $data_with_ref = $collection->toArray(null, true, true, true);

        // 获取 collection 中的列索引
        $this->getSheetTargetColumnIndex($data_with_ref);

        if ($this->_is_with_raw_data) {

            $ret = [
                'format_data' => $this->dataFormat($data_with_ref),
                'raw_data' => $this->_is_raw_data_with_cell_ref ? $data_with_ref : $collection->toArray(),

            ];


        } else {

            $ret = $this->dataFormat($data_with_ref);

        }

        return $ret;

    }

    /**
     * @param $data_with_ref
     * Desc: 获取 collection 中的列索引
     * User: Whx
     * Date: 2021-07-26 20:24
     */
    private function getSheetTargetColumnIndex($data_with_ref)
    {

        // 当前 collection 从哪一行读取
        $this->getCurrRowIndex();

        // 获取表头
        $this->_active_sheet_header = $data_with_ref[$this->_curr_row_index];

        if (!$this->_active_sheet_header) {
            throw new \Exception('The row index is out of range.please check it again.');
        }

        foreach ($this->_active_sheet_header as $key => $item) {  // 部分表格会读取空列数据
            if (!$item) {
                unset($this->_active_sheet_header[$key]);
            }
        }

        // 获取当前 collection 的目标列名
        $this->getCurrTargetColumnName();

        // 目标列名转换为对应的列索引
        $this->getCurrTargetColumnIndex();

    }

    /**
     * Desc: 获取当前 sheet 表的行号
     * User: Whx
     * Date: 2021-07-26 20:24
     */
    private function getCurrRowIndex()
    {
        if (!$this->_is_multi_sheet) { // 单个 sheet

            // 没设置，则默认为第 1 行
            $this->_curr_row_index = $this->_row_index ?: 1;

        } else { // 多个 sheet

            if (!$this->_row_index) { // 没设置，则默认为全部列

                $this->_curr_row_index = 1;

            } else {

                // 没设置对应的 sheet 列数据，则默认为全部列，否则按设置列为准
                $this->_curr_row_index = $this->_row_index[$this->_curr_collection_key] ?: 1;

            }

        }
    }

    /**
     * Desc: 获取当前 sheet 表的列名
     * User: Whx
     * Date: 2021-07-26 20:24
     */
    private function getCurrTargetColumnName()
    {
        if (!$this->_is_multi_sheet) { // 单个 sheet

            // 没设置，则默认为全部列，否则按设置列为准
            $this->_curr_sheet_target_column_name = $this->_sheet_target_column_name ?: array_values($this->_active_sheet_header);

        } else { // 多个 sheet

            if (!$this->_sheet_target_column_name) { // 没设置，则默认为全部列

                $this->_curr_sheet_target_column_name = array_values($this->_active_sheet_header);

            } else {

                // 没设置对应的 sheet 列数据，则默认为全部列，否则按设置列为准
                $this->_curr_sheet_target_column_name = $this->_sheet_target_column_name[$this->_curr_collection_key] ?: array_values($this->_active_sheet_header);

            }

        }
    }

    private function getCurrTargetColumnIndex()
    {
        // 重置列索引
        $this->_curr_sheet_target_column_index = [];

        // 获取列索引
        foreach ($this->_active_sheet_header as $k => $v) {

            if (in_array($v, $this->_curr_sheet_target_column_name)) {

                array_push($this->_curr_sheet_target_column_index, $k);

            }

        }
    }


    /**
     * @return mixed
     * Desc: 未指定工作表时，默认取第一个工作表数据
     * User: Whx
     * Date: 2021-07-23 13:56
     */
    private function getDefaultSheetData()
    {

        $this->_spreadsheet->setActiveSheetIndex($this->_default_active_sheet_index);

        $this->_sheet_collection[$this->_default_active_sheet_index] = $this->_spreadsheet->getActiveSheet();

        $this->_is_multi_sheet = false;

    }


    /**
     * @param $file_name
     * Desc: 输入文件校验
     * User: Whx
     * Date: 2021-07-26 19:59
     */
    private function inputFileCheck($file_name)
    {

        if (!file_exists($file_name)) {

            throw new \Exception(" File doesn't exist.");

        }

        $tmp = explode('.', $file_name);

        $suffix = end($tmp);

        if (!in_array($suffix, $this->_suport_suffix)) {

            throw new \Exception("Did not support this file type: " . $suffix . "【support file type is : " . implode('|', $this->_suport_suffix) . "】");

        }

        $this->_input_file = $file_name;
    }

    /**
     * @param $data_with_ref
     * Desc: Description
     * User: Whx
     * Date: 2021-07-28 17:03
     */
    public function dataFormat($data_with_ref)
    {
        $ret = [];

        // 去除表头数据
        for ($i = 0; $i <= $this->_curr_row_index; $i++) {

            unset($data_with_ref[$i]);

        }

        foreach ($data_with_ref as $k => $item) {

            $tmp = [];

            foreach ($item as $k1 => $v1) {

                if (in_array($k1, $this->_curr_sheet_target_column_index)) {

                    $column_name = $this->_active_sheet_header[$k1];

                    $tmp[$column_name] = in_array($column_name,$this->_date_column_name) ?
                        date($this->_date_format, Date::excelToTimestamp($v1)) :
                        $v1;

                }

            }

            array_push($ret, $tmp);

        }

        return $ret;


    }


}
