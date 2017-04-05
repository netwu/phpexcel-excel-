<?php
/***********************************************************************************************************
 * 导出excel文件 
 * $author:wyf1jobs@163.com
 * 1.如果在导出文件时设置边框格式会增加资源消耗, 所以建议在模板中设置格式.
 * 2.如果模板格式存在兼容性问题, 建议在程序中设置好格式然后导出用作模板, 再把格式设置代码注释.
 * 3.在模板中把相应字段的键值设置好即可
 * resume_list 数据格式见 resume.json
 ***********************************************************************************************************/
class exportExcel
{
    private $excel_template = 'excel_template.xlsx';
    private $fileDir        = 'Uploads';

    /**
     * 导出excel
     * @param  [type] $resume [description]
     * @return [type]         [description]
     */
    public function exportExcel($resume_list){
        $filePath = $this->createExcel($resume_list);

        if(false !== strpos($_SERVER['HTTP_USER_AGENT'],'Firefox')){
            $this->dl_file($filePath);
        }else{
            $outputFileName = str_replace($this->fileDir.'/', '', $filePath);
            header('Pragma:public');
            header('Content-Type:application/x-msexecl;name="'.$outputFileName.'"');
            header('Content-type: application/xls'); 
            header('cache-control:must-revalidate');                          
            header("Content-Type: application/force-download");
            header("Content-Type: application/octet-stream");
            header("Content-Transfer-Encoding: binary");    
            header('Content-Disposition:inline;filename="'.$outputFileName.'"');
            header("content-Type: text/html; charset=Utf-8"); 
            header("Content-Transfer-Encoding: utf-8");
            header("Expires: Mon, 26 Jul 1997 05:00:00 GMT");
            header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
            header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
            header('cache-control:must-revalidate');             
            header("Content-Type: application/download");
            echo file_get_contents($filePath);
        }
        unlink($filePath);  
     }

     /**
      * 创建excel文件
      * @param  array $resume_list 简历信息
      * @return string             excel文件路径
      */
    public function createExcel($resume_list){
        //加载 phpexcel
        $this->phpExcelInclude();
        $filename = 'demo.xlsx';
        $filePath = $this->fileDir.'/'.$filename;
        
        $phpExcel = \PHPExcel_IOFactory::load($this->excel_template);
        $phpExcel = $this->setIndexSheet($resume_list, $phpExcel);
        $phpExcel = $this->setDetailSheet($resume_list, $phpExcel);
        $objWriter = \PHPExcel_IOFactory::createWriter($phpExcel, 'Excel2007');

        $objWriter->save($filePath);

        return $filePath;
    }


    /**
     * 设置第一个工作表
     * @param [type] $resume_list [description]
     * @param [type] $phpExcel    [description]
     */
    private function setIndexSheet($resume_list, $phpExcel){
        //设置列数
        $columnNum = 'XVII';
        $currentSheet = $phpExcel->getSheetByName('UserList');
        $title = count($resume_list) == 1 ? $resume_list[0]['name'] : 'UserList';
        $currentSheet->setTitle($title);
        //行数,默认10行
        $rowNum = max(count($resume_list)+1, 10);
        for($colIndex='A'; $colIndex<=$columnNum; $colIndex++){
            for($rowIndex = 1; $rowIndex<=$rowNum; $rowIndex++){
                $resume = $resume_list[0]; //获取简历列表中第一个简历, 其他的简历格式跟随第一个即可
                $addr = $colIndex.$rowIndex;
                $cell = $currentSheet->getCell($addr)->getValue();
                //设置边框
                // $currentSheet->getStyle($addr)->applyFromArray($this->getBorderStyle('CDCDCD'));

                if($cell instanceof PHPExcel_RichText){//富文本转换字符
                    $cell = $cell->__toString();
                }
                if(isset($resume[$cell])){//如果简历中有这个key
                    $temIndex = $rowIndex;
                    //自动宽度
                    foreach ($resume_list as $value) { //遍历简历列表,跟随第一个简历格式把对应的值填充到下一列
                        $tempAddr = $colIndex.$temIndex;

                        $currentSheet->setCellValue($tempAddr,$value[$cell]);
                        //靠左对齐
                        // $objStyle = $currentSheet->getStyle($tempAddr);
                        // $objAlign = $objStyle->getAlignment();
                        // $objAlign->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
                        
                        $temIndex++;
                    }
                }
            }
        }
        return $phpExcel;
    }

    /**
     * 设置excel详细信息
     * @param array $resume_list 简历列表
     * @param obj $phpExcel    phpExcel操作对象
     */
    private function setDetailSheet($resume_list, $phpExcel){
        foreach ($resume_list as $resume) {
            $detailSheet = $phpExcel->getSheetByName('detail');
            $currentSheet = clone $detailSheet;
            $currentSheet->setTitle($resume['tit']);
            $phpExcel->addSheet($currentSheet);

            //设置列数
            $columnNum = 'III';
            $rowNum = $currentSheet->getHighestRow();
            for($colIndex='A';$colIndex<=$columnNum;$colIndex++){
                for($rowIndex = 1; $rowIndex<=$rowNum; $rowIndex++){
                    $addr = $colIndex.$rowIndex;
                    $objStyle = $currentSheet->getStyle($addr);
                    //设置边框 灰色
                    // $objStyle->applyFromArray($this->getBorderStyle('CDCDCD'));
                    $cell = $currentSheet->getCell($addr)->getValue();
                    if($cell instanceof PHPExcel_RichText){//富文本转换字符
                        $cell = $cell->__toString();
                    }
                    // if($cell){
                        //设置边框 蓝色
                        // $objStyle->applyFromArray($this->getBorderStyle('9ECDFD'));   
                    // }
                   
                    // 设置颜色
                    // if(false !== strpos($cell,'{bgcolor}')){
                    //     $objFill = $objStyle->getFill(); 
                    //     $objFill->setFillType(\PHPExcel_Style_Fill::FILL_SOLID); 
                    //     $objFill->getStartColor()->setARGB('9ECDFD');  
                    //     $currentSheet->setCellValue($addr,str_replace('{bgcolor}', '', $cell));
                    // }
                    
                    if(isset($resume[$cell])){//如果简历中有这个key
                        $currentSheet->setCellValue($addr,$resume[$cell]);
                        //识别格式
                        $objStyle->getAlignment()->setWrapText(true);
                        //自适应行高
                        $currentSheet->getRowDimension($rowIndex)->setRowHeight(-1);
                        //设置行宽
                        $currentSheet->getColumnDimension($colIndex)->setWidth(60);  
                    }
                }
            }
        }
        //过河拆桥,删除默认的sheet模板
        $phpExcel->removeSheetByIndex(1);
        return $phpExcel;
    }

    /**
     * 获取excel 边框样式
     * @return [type] [description]
     */
    private function getBorderStyle($color){
        return array(
         'borders' => array (
               'outline' => array (
                     'style' => \PHPExcel_Style_Border::BORDER_THIN,   //设置border样式
                     //'style' => \PHPExcel_Style_Border::BORDER_THICK,  另一种样式
                     'color' => array ('argb' => $color),          //设置border颜色
              ),
        ),
      );
    }

    
    /**
     *#引进相关文件 phpexcel
     */
    public function phpExcelInclude(){
        Vendor("phpExcel.Classes.PHPExcel");
        Vendor("phpExcel.Classes.PHPExcel.IOFactory");
    }

    /**
     * 下载
     * @param  [type] $filePath [description]
     * @return [type]           [description]
     */
     public function download($filePath){
        $filename = str_replace($this->fileDir.'/', '', $filePath);

        header("Content-Type: application/force-download");
        header("Content-type: text/html; charset=utf-8");

        header("Content-Transfer-Encoding: binary");
        header('Content-Type: application/zip');
        header('Content-Disposition: attachment; filename='.$filename);
        header("Connection: close");
        readfile($filePath);
    }
    
     /**
      * 兼容火狐浏览器下载方式
      * @param  [type] $file [description]
      * @return [type]       [description]
      */
     function dl_file($filePath){
        //First, see if the file exists
        if (!is_file($filePath)) { die("<b>404 File not found!</b>"); }
         
            //Gather relevent info about file
            $len = filesize($filePath);
            $filename = str_replace($this->fileDir.'/', '', $filePath);
            $file_extension = strtolower(substr(strrchr($filename,"."),1));
            //This will set the Content-Type to the appropriate setting for the file
            switch( $file_extension ) {
              case "pdf": $ctype="application/pdf"; break;
              case "exe": $ctype="application/octet-stream"; break;
              case "zip": $ctype="application/zip"; break;
              case "doc": $ctype="application/msword"; break;
              case "xls": $ctype="application/vnd.ms-excel"; break;
              case "xlsx": $ctype="application/vnd.ms-excel"; break;
              case "ppt": $ctype="application/vnd.ms-powerpoint"; break;
              case "gif": $ctype="image/gif"; break;
              case "png": $ctype="image/png"; break;
              case "jpeg":
              case "jpg": $ctype="image/jpg"; break;
              case "mp3": $ctype="audio/mpeg"; break;
              case "wav": $ctype="audio/x-wav"; break;
              case "mpeg":
              case "mpg":
              case "mpe": $ctype="video/mpeg"; break;
              case "mov": $ctype="video/quicktime"; break;
              case "avi": $ctype="video/x-msvideo"; break;
         
              //The following are for extensions that shouldn't be downloaded (sensitive stuff, like php files)
              case "php":
              case "htm":
              case "html":
              case "txt": die("<b>Cannot be used for ". $file_extension ." files!</b>"); break;
         
              default: $ctype="application/force-download";
        }
        //Begin writing headers
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Cache-Control: public"); 
        header("Content-Description: File Transfer");
         
        //Use the switch-generated Content-Type
        header("Content-Type: $ctype");
     
        //Force the download
        $header="Content-Disposition: attachment; filename=".$filename.";";
        header($header );
        header("Content-Transfer-Encoding: binary");
        header("Content-Length: ".$len);
        @readfile($filePath);
        exit;
    }

}