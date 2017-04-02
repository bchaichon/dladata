<meta charset="utf-8">
<?php

/** PHPExcel */
require_once 'Classes/PHPExcel.php';

/** PHPExcel_IOFactory - Reader */
include 'Classes/PHPExcel/IOFactory.php';


$inputFileName = "idlocal.xls";  
$inputFileType = PHPExcel_IOFactory::identify($inputFileName);  
$objReader = PHPExcel_IOFactory::createReader($inputFileType);  
$objReader->setReadDataOnly(true);  
$objPHPExcel = $objReader->load($inputFileName);  



$objWorksheet = $objPHPExcel->setActiveSheetIndex(0);
$highestRow = $objWorksheet->getHighestRow();
$highestColumn = $objWorksheet->getHighestColumn();

$headingsArray = $objWorksheet->rangeToArray('A1:'.$highestColumn.'1',null, true, true, true);
$headingsArray = $headingsArray[1];

$r = -1;
$namedDataArray = array();
for ($row = 2; $row <= $highestRow; ++$row) {
    $dataRow = $objWorksheet->rangeToArray('A'.$row.':'.$highestColumn.$row,null, true, true, true);
    if ((isset($dataRow[$row]['A'])) && ($dataRow[$row]['A'] > '')) {
        ++$r;
        foreach($headingsArray as $columnKey => $columnHeading) {
                $data=$dataRow[$row][$columnKey];
                //echo $columnKey ; echo ":";
                if($columnKey=="D" || $columnKey=="E" || $columnKey=="J"){
                    $data = PHPExcel_Style_NumberFormat::toFormattedString($data, "YYYY-M-D"); //2016-01-29
                }
                //

                $namedDataArray[$r][$columnHeading] = $data;
          
        }
        //echo "<br>";
    }
}


$count=0;
foreach ($namedDataArray as $result) {
$count++;

        
        $tmp = explode("ตำบล", $result["dla_name"]) ;

        if(sizeof($tmp)==2){
           $tambon = $tmp[1];
           $typename= $tmp[0]."ตำบล";
        }else{
           $tmp2 = explode("เมือง", $result["dla_name"]) ; 
           
          if(sizeof($tmp2)==2){
           $tambon = $tmp2[1];
           $typename= $tmp2[0]."เมือง"; 
          }else{

            $tmp3 = explode("จังหวัด", $result["dla_name"]) ; 
           if(sizeof($tmp3)==2){
             $tambon = "ในเมือง";
             $typename= $tmp3[0]."จังหวัด"; 
            }
          }
        }

        $strSQL = "";
        $strSQL .= "INSERT INTO dladata ";
        $strSQL .= "(dla_id,dla_name,dla_province,dla_amphur,dla_tambon , dla_typename ) ";
        $strSQL .= "VALUES ";
        $strSQL .= "('".$result["dla_id"]."','".$result["dla_name"]."' ";
        $strSQL .= ",'".$result["dla_province"]."','".$result["dla_amphur"]."' , '".$tambon."' , '".$typename."' ";
       

       
        $strSQL .= "  ) ; ";
        
        echo $strSQL ; echo "<br>";
        $query = mysql_query($strSQL);
        $in="";
        if($query){
          $in="insert : $count";
        }
}
?>
