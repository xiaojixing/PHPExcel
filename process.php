<?php
header("Content-type:text/html;charset=utf-8");
require_once 'PHPExcel/Classes/PHPExcel.php';
require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';

class HandleData
{

    public $mysql_conf = [];
    public $file;
    public $mysqli;
    public $cellName;

    public function __construct($mysql_config, $file)
    {
        $this->mysql_conf = $mysql_config;
        $this->file       = $file;
        $this->cellName   = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ');
        $this->connectMysql();
        $this->filterFormat();
    }

    private function connectMysql()
    {
        $this->mysqli = @new mysqli($this->mysql_conf['host'], $this->mysql_conf['db_user'], $this->mysql_conf['db_pwd']);
        if ($this->mysqli->connect_errno) {
            die("连接数据库失败:\n" . $this->mysqli->connect_error); //诊断连接错误
        }
        $this->mysqli->query("set names 'utf8';"); //编码转化
        $select_db = $this->mysqli->select_db($this->mysql_conf['db']);
        if (!$select_db) {
            die("could not connect to the db:\n" . $this->mysqli->error);
        }
    }

    public function filterFormat()
    {
        //判断是否选择了要上传的表格
        if (empty($this->file)) {
            $res['ok']            = 0;
            $res['error_message'] = '您未选择表格';
        }

        //获取表格的大小，限制上传表格的大小5M
        $file_size = $this->file['size'];
        if ($file_size > 5 * 1024 * 1024) {
            $res['ok']            = 0;
            $res['error_message'] = '上传失败，上传的表格不能超过5M的大小';
        }
        //限制上传表格类型
        $file_type = $_FILES['myfile']['type'];
        //application/vnd.ms-excel  为xls文件类型
        $authorized = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel',
        ];

        if (!in_array($file_type, $authorized)) {
            $res['ok']            = 0;
            $res['error_message'] = '上传失败，上传的表格类型不合法';
        }

    }
 
/**
 *  数据导入
 * @param string $file excel文件
 * @param string $sheet
 * @return string   返回解析数据
 * @throws PHPExcel_Exception
 * @throws PHPExcel_Reader_Exception
 */
    public function importExecl($sheet = 0)
    {

        $this->file['tmp_name'] = iconv("utf-8", "gb2312", $this->file['tmp_name']); //转码
        if (empty($this->file['tmp_name']) or !file_exists($this->file['tmp_name'])) {
            die('file not exists!');
        }

        $objRead = new PHPExcel_Reader_Excel2007(); //建立reader对象
        if (!$objRead->canRead($this->file['tmp_name'])) {
            $objRead = new PHPExcel_Reader_Excel5();
            if (!$objRead->canRead($this->file['tmp_name'])) {
                die('No Excel!');
            }
        }

        $obj       = $objRead->load($this->file['tmp_name']); //建立excel对象
        $currSheet = $obj->getSheet($sheet); //获取指定的sheet表
        $columnH   = $currSheet->getHighestColumn(); //取得最大的列号
        $columnCnt = array_search($columnH, $this->cellName);
        $rowCnt    = $currSheet->getHighestRow(); //获取总行数

        $data = array();
        for ($_row = 1; $_row <= $rowCnt; $_row++) {
            //读取内容
            for ($_column = 0; $_column <= $columnCnt; $_column++) {
                $cellId    = $this->cellName[$_column] . $_row;
                $cellValue = $currSheet->getCell($cellId)->getValue();
                //$cellValue = $currSheet->getCell($cellId)->getCalculatedValue();  #获取公式计算的值
                if ($cellValue instanceof PHPExcel_RichText) {
                    //富文本转换字符串
                    $cellValue = $cellValue->__toString();
                }

                $data[$_row][$this->cellName[$_column]] = $cellValue;
            }
        }
        
        return $data;
    }

/**
 * 数据导出
 * @param array $title   标题行名称
 * @param array $data   导出数据
 * @param string $fileName 文件名
 * @param string $savePath 保存路径
 * @param $type   是否下载  false--保存   true--下载
 * @return string   返回文件全路径
 * @throws PHPExcel_Exception
 * @throws PHPExcel_Reader_Exception
 */
    public function exportExcel($title = array(), $data = array(), $fileName = '', $savePath = './', $isDown = false)
    {

        $obj = new PHPExcel();

        //横向单元格标识
        $obj->getActiveSheet(0)->setTitle('sheet名称'); //设置sheet名称
        $_row = 1; //设置纵向单元格标识
        if ($title) {
            $_cnt = count($title);
            $obj->getActiveSheet(0)->mergeCells('A' . $_row . ':' . $this->cellName[$_cnt - 1] . $_row); //合并单元格
            $obj->setActiveSheetIndex(0)->setCellValue('A' . $_row, '数据导出：' . date('Y-m-d H:i:s')); //设置合并后的单元格内容
            $_row++;
            $i = 0;
            foreach ($title as $v) {
                //设置列标题
                $obj->setActiveSheetIndex(0)->setCellValue($this->cellName[$i] . $_row, $v);
                $i++;
            }
            $_row++;
        }

        //填写数据
        if ($data) {
            $i = 0;
            foreach ($data as $_v) {
                $j = 0;
                foreach ($_v as $_cell) {
                    $obj->getActiveSheet(0)->setCellValue($this->cellName[$j] . ($i + $_row), $_cell);
                    $j++;
                }
                $i++;
            }
        }

        //文件名处理
        if (!$fileName) {
            $fileName = uniqid(time(), true);
        }

        $objWrite = PHPExcel_IOFactory::createWriter($obj, 'Excel2007');

        if ($isDown) {
            //网页下载
            header('pragma:public');
            header("Content-Disposition:attachment;filename=$fileName.xls");
            $objWrite->save('php://output');exit;
        }

        $_fileName = iconv("utf-8", "gb2312", $fileName); //转码
        $_savePath = $savePath . $_fileName . '.xlsx';
        $objWrite->save($_savePath);
        return $savePath . $fileName . '.xlsx';
    }

}

$conf = array(
    'host'     => '127.0.0.1',
    'db'       => 'excel',
    'db_table' => '',
    'db_user'  => 'root',
    'db_pwd'   => '',
);

$file   = $_FILES['myfile'];
$handle = new HandleData($conf, $file);
//$handle->excelToArray();
$handle->importExecl();
//$handle->exportExcel(array('姓名', '年龄'), array(array('a', 21), array('b', 23)), '档案', './', true);
