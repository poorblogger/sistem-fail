<html>
  <head>
    <style type="text/css">
    table {
        border-collapse: collapse;
    }        
    td {
        border: 1px solid black;
        padding: 0 0.5em;
    }        
    </style>
  </head>
  <body>
  <title>Sistem Carian Fail LPM</title>
    <table>
    <h3><center>SISTEM CARIAN FAIL LPM</center></h3>
     <div>
           <center> <img src="image/logo.png" alt="myPic" /></center>
        </div>
   <h1><font color="red"><center> MAKLUMAT CARIAN</center></font></h1>
   

    </table>
  </body>
</html>


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


$search_var=strtoupper($_POST['search']);
//$search_var=($_POST['search']);

$kira= 0;

for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {
 
  
 

  if( preg_match("/" . $search_var . "/", $data->sheets[0]['cells'][$i][4])  ) {
    echo "<center><td>";
    echo "Fail   ";
    echo $data->sheets[0]['cells'][$i][4];
    echo " No Rujukan :  ";
    echo $data->sheets[0]['cells'][$i][3];
     echo " Di Lokasi  ";
    echo $data->sheets[0]['cells'][$i][2];
    
    echo "</td></center>";
    $kira=$kira+1;

    
   }//if preg_match
   

}//end for



  
  if ( $kira==0 )
    {
        echo "<h2><font color='blue'><center>Carian tiada </center></h2></font>";
         echo "<center> <a href='index.php'><image src='image/back.jpg' height='40'></a></center>";

    }
  
  else
  {
    echo "<center>";
    
    echo "<br/>Jumlah Fail Ditemui<br/>";
    echo $kira;
    echo"</center>";
    echo "<center> <a href='index.php'><image src='image/back.jpg' height='40'></a></center>";
  }
  

?>
