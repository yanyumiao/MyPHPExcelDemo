<?php
require_once '/var/www/PHPExcel-1.8.1/Classes/PHPExcel.php';
//$inputName="2007.xlsx";
$inputName="2003.xls";
$objPHPExcel = PHPExcel_IOFactory::load($inputName);

// ------------------------------------- Read ----------------------------------------
// read to array
$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
print_r($sheetData);
$sheetData = $objPHPExcel->getActiveSheet()->toArray(); // default: null,true,true,false
print_r($sheetData);

// read cell
$cellValue = $objPHPExcel->getActiveSheet()->getCell('A1')->getValue();
var_dump($cellValue);
$cellValue = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(0,1)->getValue(); // Notice: column from 0, row from 1
var_dump($cellValue);

// ------------------------------------- Write ----------------------------------------
$objPHPExcel->getActiveSheet()->setCellValue('A1', "hello");
$cellValue = $objPHPExcel->getActiveSheet()->getCell('A1')->getValue();
var_dump($cellValue);
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,1,null); // Notice: column from 0, row from 1
$cellValue = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(1,1)->getValue();
var_dump($cellValue);

// ------------------------------------- Delete ----------------------------------------
$objPHPExcel->getActiveSheet()->removeRow(1,1); //Notice: from row 1, total 1 row & the row number n move to row number n-1
$sheetData = $objPHPExcel->getActiveSheet()->toArray();
print_r($sheetData);
// or you can use setCellVlue() or setCellValueByColumnAndRow() to set null
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(0,2,null);
$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,2,null);
$sheetData = $objPHPExcel->getActiveSheet()->toArray();
print_r($sheetData);

// ------------------------------------- Save ----------------------------------------
//$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
//$objWriter->save($inputName);


// 暂不推荐使用下面的方法
// 测试发现对于xls类型的excel，在使用PHPExcel_Writer_Excel5保存后会出错
// ------------------------------------- More ----------------------------------------
//$inputFileType = PHPExcel_IOFactory::identify($inputName);  
//$objReader = PHPExcel_IOFactory::createReader(inputFileType);
//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $inputFileType);

// ------------------------------------- Error ----------------------------------------
//$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
//$objWriter->save($inputName);
