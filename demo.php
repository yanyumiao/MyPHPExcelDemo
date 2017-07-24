<?php
require_once '/var/www/PHPExcel-1.8.1/Classes/PHPExcel.php';
//$inputName="demo.xlsx";
$inputName="2003.xls";
$objPHPExcel = PHPExcel_IOFactory::load($inputName);

// ------------------------------------- Read ----------------------------------------
// read to array
$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
print_r($sheetData);
$sheetData = $objPHPExcel->getActiveSheet()->toArray(); // default: null,true,true,false
print_r($sheetData);

// read cell
//$cellValue = $objPHPExcel->getActiveSheet()->getCell('A1')->getValue();
$cellValue = $objPHPExcel->getActiveSheet()->getCellByColumnAndRow(0,1)->getValue(); // Notice: column from 0, row from 1
var_dump($cellValue);

// ------------------------------------- Write ----------------------------------------
//$objPHPExcel->getActiveSheet()->setCellValue('A1', "hello");
//$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow(1,1,null); // Notice: column from 0, row from 1

// ------------------------------------- Delete ----------------------------------------
//$objPHPExcel->getActiveSheet()->removeRow(1,1); //Notice: from row 1, total 1 row & the row number n move to row number n-1
// or you can use setCellVlue() or setCellValueByColumnAndRow() to set null 
$sheetData = $objPHPExcel->getActiveSheet()->toArray();
print_r($sheetData);
//exit;

// ------------------------------------- Save ----------------------------------------
//$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
//$objWriter->save($inputName);

// ------------------------------------- Bug ----------------------------------------
// Bug with Excel5 
//$inputFileType = PHPExcel_IOFactory::identify($inputName);  
//$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $inputFileType); // error when read agin
//$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel); // pass
//$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel); // error when read agin
//$objWriter->save($inputName);


