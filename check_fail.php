<?php
// Test CVS

require_once 'Excel/reader.php';


// ExcelFile($filename, $encoding);
$data = new Spreadsheet_Excel_Reader();


// Set output Encoding.
$data->setOutputEncoding('CP1251');

/***
* if you want you can change 'iconv' to mb_convert_encoding:
* $data->setUTFEncoder('mb');
*
**/

/***
* By default rows & cols indeces start with 1
* For change initial index use:
 $data->setRowColOffset(0);
*
**/



/***
*  Some function for formatting output.
 $data->setDefaultFormat('%.4f');
 /*
* setDefaultFormat - set format for columns with unknown formatting
*/
$data->setColumnFormat(4, '%.3f');
/* setColumnFormat - set format for column (apply only to number fields)
*
**/

$data->read('fail.xls');

/*/


 $data->sheets[0]['numRows'] - count rows
 $data->sheets[0]['numCols'] - count columns
 $data->sheets[0]['cells'][$i][$j] - data from $i-row $j-column

 $data->sheets[0]['cellsInfo'][$i][$j] - extended info about cell
    
    $data->sheets[0]['cellsInfo'][$i][$j]['type'] = "date" | "number" | "unknown"
        if 'type' == "unknown" - use 'raw' value, because  cell contain value with format '0.00';
    $data->sheets[0]['cellsInfo'][$i][$j]['raw'] = value if cell without format 
    $data->sheets[0]['cellsInfo'][$i][$j]['colspan'] 
    $data->sheets[0]['cellsInfo'][$i][$j]['rowspan'] 
*/

     

error_reporting(E_ALL ^ E_NOTICE);

$search_var=$_POST['search'];
$kira= 0;

for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {

  if( preg_match("/" . $search_var . "/", $data->sheets[0]['cells'][$i][4]) ) {
    echo "<br/>";
    echo "Fail ";
    echo $data->sheets[0]['cells'][$i][4];
    echo " Berada di Level ";
    echo $data->sheets[0]['cells'][$i][2];
    echo"";
    $kira=$kira+1;
    
   }//if preg_match
   

}//end for


  
  if ( $kira==0 )
    {
        echo "Carian tiada sila <a href='index.php'>Klik Untuk Kembali</a>";

    }
  
  
  

?>
