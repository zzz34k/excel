<?php

/**
 * MS Excel 操作类:用于Excel的生成和下载
 * 依赖 PHPExcel控件, Yii框架
 * @author yangaimin
 *
 */
class Excel
{
    private $_excel;

    // private $_columns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    private $_columns = array(
        'A',
        'B',
        'C',
        'D',
        'E',
        'F',
        'G',
        'H',
        'I',
        'J',
        'K',
        'L',
        'M',
        'N',
        'O',
        'P',
        'Q',
        'R',
        'S',
        'T',
        'U',
        'V',
        'W',
        'X',
        'Y',
        'Z',
        'AA',
        'AB',
        'AC',
        'AD',
        'AE',
        'AF',
        'AG',
        'AH',
        'AI',
        'AJ',
        'AK',
        'AL',
        'AM',
        'AN',
        'AO',
        'AP',
        'AQ',
        'AR',
        'AS',
        'AT',
        'AU',
        'AV',
        'AW',
        'AX',
        'AY',
        'AZ',
        'BA',
        'BB',
        'BC',
        'BD',
        'BE',
        'BF',
        'BG',
        'BH',
        'BI',
        'BJ',
        'BK',
        'BL',
        'BM',
        'BN',
        'BO',
        'BP',
        'BQ',
        'BR',
        'BS',
        'BT',
        'BU',
        'BV',
        'BW',
        'BX',
        'BY',
        'BZ',
        'CA',
        'CB',
        'CC',
        'CD',
        'CE',
        'CF',
        'CG',
        'CH',
        'CI',
        'CJ',
        'CK',
        'CL',
        'CM',
        'CN',
        'CO',
    );
    private $_version = "Excel5";
    private $_isNew;

    /**
     * Excel constructor.
     * @param string $version
     * @param bool $isNew
     * @param string $filename
     */
    public function Excel($version = "Excel5", $isNew = true, $filename = '')
    {
        $this->_version = $version;
        $this->_isNew = $isNew;
        if ($this->_isNew) {
            $this->_excel = new PHPExcel();
            $this->_excel->getProperties()->setCreator("aimin.yang@gmail.com"); // 创建人
            $this->_excel->getProperties()->setLastModifiedBy("aimin.yang@gmail.com"); // 最后修改人
            $this->_excel->getProperties()->setTitle("Excel generator"); // 标题
            $this->_excel->getProperties()->setSubject("Report Export"); // 题目
            $this->_excel->getProperties()->setDescription("Report export samples"); // 描述
            $this->_excel->getProperties()->setKeywords("phpexcel yii "); // 关键字
            $this->_excel->getProperties()->setCategory("report"); // 种类
        } else {
            $finalFileName = Yii::app()->basePath . '/runtime/' . $filename;
            $this->_excel = PHPExcel_IOFactory::load($finalFileName);
        }


        // $this->_columns = str_split($this->_columns);
    }

    public function __destruct()
    {
        $this->_excel->disconnectWorksheets();
        PHPExcel_Calculation::unsetInstance($this->_excel);
    }

    /**
	 * Create Excel sheet
     * @param $index
     * @param $title
     * @param $dataSet
     * @param array $header
     * @param array $default
     * @param array $extHeader
     * @return $this
     * @throws Exception
     */
    public function createSheet($index, $title, $dataSet, $header = array(),
                                $default = array(), $extHeader = array())
    {
        if ($this->_isNew) {
            if (is_array($dataSet)) {
                $this->createFromArray($index, $title, $dataSet, $header, $default, $extHeader);
            } elseif (is_object($dataSet)) {
                $this->createFromDataProvider($index, $title, $dataSet, $header, $default, $extHeader);
            } else {
                throw new Exception('请输入支持的数据集类型');
            }
        } else {
            if (is_array($dataSet)) {
                $this->addDataToExcel($index, $title, $dataSet, $header, $default, $extHeader);
            } elseif (is_object($dataSet)) {
                $this->addDataProviderToExcel($index, $title, $dataSet, $header, $default, $extHeader);
            } else {
                throw new Exception('请输入支持的数据集类型');
            }
        }

        return $this;
    }

    /**
     * create excel Sheet from array
     *
     * @param int $index
     * @param string $title
     * @param mixed $dataSet
     * @param string $header
     */
    protected function createFromArray($index, $title, $dataSet, $header = array(),
                                       $default = array(), $extHeader = array())
    {
        $this->_excel->setActiveSheetIndex($index);
        $sheet = $this->_excel->getActiveSheet();
        $sheet->setTitle($title);

        $rowIndex = 1;
        $colIndex = 0;

        if (!empty($extHeader)) {
            foreach ($extHeader as $k => $v) {
                $value = isset($v['value']) ? $v['value'] : $v;
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($value));
            }

            foreach ($extHeader as $k => $v) {
                if (isset($v['colspan']) && $v['colspan'] > 1) {
                    $sheet->mergeCells($this->_columns[$v['start']] . "1:" . $this->_columns[$v['end']] . "1");
                }
            }
            $rowIndex++;
        }

        $colIndex = 0;

        if (!empty($header)) {
            foreach ($header as $k => $v) {
//                echo $k."  ".$v.PHP_EOL;
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($v));
            }
            $rowIndex++;
        }
//        $objActSheet->mergeCells('B1:C22');
        $merge = array();
        if (!empty($dataSet)) {
            foreach ($dataSet as $data) {
                $colIndex = 0;
                if (empty($header)) {
                    foreach ($data as $k => $v) {
                        $val = !isset($v) ? (isset($default[$k]) ? $default[$k] : '') : $v;
                        $val = trim(str_replace("<br>", "\n", $val));
                        $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
                    }
                } else {
                    foreach ($header as $k => $v) {
                        $val = isset($data[$k]) ? $data[$k] : (isset($default[$k]) ? $default[$k] : '');

//                        error_log(var_export($val, true) . "\r\n", 3, 'dev.log');
                        if (is_string($val)) {
                            if (strpos($val, 'needSpan') !== false) {
                                $mergeRow = str_replace('needSpan', '', substr($val, strpos($val, 'needSpan')));
                                $merge[] = $this->_columns[$colIndex] . $rowIndex . ":" . $this->_columns[$colIndex] . ($rowIndex + $mergeRow - 1);
                                $val = substr($val, 0, strpos($val, 'needSpan'));
                            }
                            $val = trim(str_replace("<br>", "\n", $val));
                        }
                        $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
                    }
                }
                $rowIndex++;
            }
        }
        if ($merge) {
            foreach ($merge as $item) {
                $sheet->mergeCells($item);
            }
        }
        $this->_excel->createSheet();
    }

    /**
     * append to Excel from array
     * @param $index
     * @param $title
     * @param $dataSet
     * @param array $header
     * @param array $default
     * @param array $extHeader
     * @params array $skipRow
     * @return $this
     * @throws PHPExcel_Exception
     */
    public function addDataToExcel($index, $title, $dataSet, $header = array(),
                                   $default = array(), $extHeader = array(), $skipRow = 1)
    {

        $this->_excel->setActiveSheetIndex($index);
        $sheet = $this->_excel->getActiveSheet();

        $rowIndex = $sheet->getHighestRow() + $skipRow;
        if (!empty($dataSet)) {
            foreach ($dataSet as $data) {
                $colIndex = 0;
                if (empty($header)) {
                    foreach ($data as $k => $v) {
                        $val = !isset($v) ? (isset($default[$k]) ? $default[$k] : '') : $v;
                        $val = trim(str_replace("<br>", "\n", $val));
                        $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
                    }
                } else {
                    foreach ($header as $k => $v) {
                        $val = isset($data[$k]) ? $data[$k] : (isset($default[$k]) ? $default[$k] : '');
//                        error_log(var_export($val, true) . "\r\n", 3, 'dev.log');
                        if (is_string($val)) {
                            $val = trim(str_replace("<br>", "\n", $val));
                        }
                        $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
                    }
                }
                $rowIndex++;
            }
        }

        return $this;
    }


    /**
     * append to Excel from dataset
     * @param $index
     * @param $title
     * @param $dataSet
     * @param array $header
     * @param array $default
     * @param array $extHeader
     * @params array $skipRow
     * @return $this
     * @throws PHPExcel_Exception
     */
    public function addDataToExcelIncHeader($index, $title, $dataSet, $header = array(),
                                   $default = array(), $extHeader = array(), $skipRow = 1)
    {

        $this->_excel->setActiveSheetIndex($index);
        $sheet = $this->_excel->getActiveSheet();

        $rowIndex = $sheet->getHighestRow() + $skipRow;

        $colIndex = 0;
        if (!empty($extHeader)) {
            foreach ($extHeader as $k => $v) {
                $value = isset($v['value']) ? $v['value'] : $v;
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($value));
            }

            foreach ($extHeader as $k => $v) {
                if (isset($v['colspan']) && $v['colspan'] > 1) {
                    $sheet->mergeCells($this->_columns[$v['start']] . $rowIndex . ":" . $this->_columns[$v['end']] .$rowIndex);
                }
            }
            $rowIndex++;
        }

        $colIndex = 0;
        if (!empty($header)) {
            foreach ($header as $k => $v) {
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($v));
            }
            $rowIndex++;
        }

        if (!empty($dataSet)) {
            foreach ($dataSet as $data) {
                $colIndex = 0;
                if (empty($header)) {
                    foreach ($data as $k => $v) {
                        $val = !isset($v) ? (isset($default[$k]) ? $default[$k] : '') : $v;
                        $val = trim(str_replace("<br>", "\n", $val));
                        $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
                    }
                } else {
                    foreach ($header as $k => $v) {
                        $val = isset($data[$k]) ? $data[$k] : (isset($default[$k]) ? $default[$k] : '');
//                        error_log(var_export($val, true) . "\r\n", 3, 'dev.log');
                        if (is_string($val)) {
                            $val = trim(str_replace("<br>", "\n", $val));
                        }
                        $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
                    }
                }
                $rowIndex++;
            }
        }

        return $this;
    }

    /**
     * create Excel from DataProvider
     *
     * @param int $index
     * @param string $title
     * @param mixed $dataSet
     * @param string $header
     */
    protected function createFromDataProvider($index, $title, $dataProvider, $header = array(),
                                              $default = array(), $extHeader = array())
    {
        $this->_excel->setActiveSheetIndex($index);
        $sheet = $this->_excel->getActiveSheet();
        $sheet->setTitle($title);
        if (empty($header)) {
            throw new Exception("标题定义不能为空");
        }

        $rowIndex = 1;
        $colIndex = 0;

        if (!empty($extHeader)) {
            foreach ($extHeader as $k => $v) {
                $value = isset($v['value']) ? $v['value'] : $v;
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($value));
            }

            foreach ($extHeader as $k => $v) {
                if (isset($v['colspan']) && $v['colspan'] > 1) {
                    $sheet->mergeCells($this->_columns[$v['start']] . "1:" . $this->_columns[$v['end']] . "1");
                }
            }
            $rowIndex++;
        }

        $colIndex = 0;
        foreach ($header as $k => $v) {
            $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($v));
        }
        $rowIndex++;

        $dataSet = $dataProvider->getData();
        $rows = count($dataSet);
        for ($i = 0; $i < $rows; $i++) {
            $colIndex = 0;
            $data = $dataProvider->data[$i];
            foreach ($header as $k => $v) {
                $value = $this->_getValue($data, $k);
                $val = isset($value) ? $value : (isset($default[$k]) ? $default[$k] : '');
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
            }
            $rowIndex++;
        }

        $this->_excel->createSheet();
    }

    /**
     * 根据dataprovider 增加数据
     * @param $index
     * @param $title
     * @param $dataProvider
     * @param array $header
     * @param array $default
     * @param array $extHeader
     * @throws Exception
     * @throws PHPExcel_Exception
     */
    protected function addDataProviderToExcel($index, $title, $dataProvider, $header = array(),
                                              $default = array(), $extHeader = array())
    {
        $this->_excel->setActiveSheetIndex($index);
        $sheet = $this->_excel->getActiveSheet();
        $rowIndex = $sheet->getHighestRow() + 1;
        $dataSet = $dataProvider->getData();
        $rows = count($dataSet);
        for ($i = 0; $i < $rows; $i++) {
            $colIndex = 0;
            $data = $dataProvider->data[$i];
            foreach ($header as $k => $v) {
                $value = $this->_getValue($data, $k);
                $val = isset($value) ? $value : (isset($default[$k]) ? $default[$k] : '');
                $sheet->setCellValue($this->_columns[$colIndex++] . $rowIndex, $this->_utf8($val));
            }
            $rowIndex++;
        }
    }

    /**
     * @param string $filename
     * @param string $dest
     * @return string
     * @throws PHPExcel_Reader_Exception
     */
    public function output($filename = 'excel.xls', $dest = '302')
    {
        $objWriter = PHPExcel_IOFactory::createWriter($this->_excel, $this->_version);
        $saveFilename = Yii::app()->basePath . '/runtime/' .
            (strtoupper(PHP_OS) == 'WINNT' ? mb_convert_encoding($filename, 'GBK', 'UTF-8') : $filename);
        $fullFilename = Yii::app()->basePath . '/runtime/' . $filename;
        $objWriter->save($saveFilename);
        if ($dest == 'http') {
            \application\components\DownloadUtil::output($fullFilename);
        } elseif ($dest == 'file') {
            return $fullFilename;
        } elseif ($dest == '302') {
            \application\components\DownloadUtil::output($fullFilename);
        }

//        if ($dest == 'http') {
//            header("Pragma: public");
//            header("Expires: 0");
//            header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
//            header("Content-Type:application/force-download");
//            header("Content-Type:application/vnd.ms-execl");
//            header("Content-Type:application/octet-stream");
//            header("Content-Type:application/download");;
//            header('Content-Disposition:attachment;filename="' . $filename . '"');
//            header("Content-Transfer-Encoding:binary");
//
//            $objWriter = PHPExcel_IOFactory::createWriter($this->_excel, $this->_version);
//            $finalFileName = Yii::app()->basePath . '/runtime/' . $filename;
//            $objWriter->save($finalFileName);
//            echo file_get_contents($finalFileName);
//        } elseif ($dest == 'file') {
//            $objWriter = PHPExcel_IOFactory::createWriter($this->_excel, $this->_version);
//            $finalFileName = Yii::app()->basePath . '/runtime/' . $filename;
//            $objWriter->save($finalFileName);
//            return $finalFileName;
//        } elseif ($dest == '302') {
//            $objWriter = PHPExcel_IOFactory::createWriter($this->_excel, $this->_version);
//            $finalFileName = Yii::app()->basePath . '/../download/' .
//                (PHP_OS == 'WINNT' ? iconv('utf-8', "gb2312", $filename) : $filename);
//            $objWriter->save($finalFileName);
//            header("location:" . Yii::app()->baseUrl . "/download/" . $filename);
//        }
    }

    /**
     * to string
     * @param unknown $value
     * @return unknown|string
     */
    protected function _toString($value)
    {
        if (is_string($value)) {
            return $value;
        } elseif (is_array($value)) {
            return join(",", $value);
        } elseif (is_object($value)) {
            return '';
        } else {
            return $value;
        }
    }

    /**
     * GBK to UTF8
     *
     * @param string $str
     * @return string
     */
    protected function _utf8($str)
    {
        // return iconv('gbk', 'utf-8', $str);
        return strip_tags($this->_toString($str));
    }

    /**
     * 根据key的类型反射调用相应的属性或方法
     *
     * @param unknown $model
     * @param unknown $key
     * @return mixed
     */
    protected function _getValue($model, $key)
    {
        if (strpos($key, "fn_") !== false) { // 方法名
            $options = explode("_", $key);
            $function = $options[1];
            if (isset($options[2])) {
                $params = $options[2];
                return $model->$function($params);
            } else {
                return $model->$function();
            }
        } elseif (strpos($key, "obj_") !== false) { // 对象
            $options = explode("_", $key);
            $object = $options[1];
            $property = $options[2];
            if (strpos($property, "fn") !== false) {
                $function = str_replace("fn", "", $property);
                return $model->$object->$function();
            } else {
                return $model->$object->$property;
            }
        } else {
            return $model->$key;
        }
    }

    /**
     * 注册导出Excel要用到的JS和CSS
     *
     * @param unknown $page
     */
    public static function RegisterScript($page, $option = "")
    {
        //Yii::app()->clientScript->registerScriptFile(Yii::app()->baseUrl . '/js/fancybox/jquery.fancybox.pack.js');
        //Yii::app()->clientScript->registerCssFile(Yii::app()->baseUrl . '/js/fancybox/jquery.fancybox.css');
        //Yii::app()->clientScript->registerScriptFile(Yii::app()->baseUrl . '/js/excel.js');
        //echo '<div class="pageuri" id="' . $page . '" option="' . $option . '"></div>';
    }

    /**
     * @return PHPExcel
     */
    public function getExcel()
    {
        return $this->_excel;
    }

    /**
     * @return array
     */
    public function getColumns()
    {
        return $this->_columns;
    }
}
?>
