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

for ($i = 1; $i <= $data->sheets[0]['numRows']; $i++) {

  if( preg_match("/" . $search_var . "/", $data->sheets[0]['cells'][$i][4]) ) {
    echo "<h3>";
    echo "Fail </h3>";
    echo $data->sheets[0]['cells'][$i][4];
    echo " Berada di Level ";
    echo $data->sheets[0]['cells'][$i][2];
    echo"";
    //echo "<br/>\n"; //for ($j = 1; $j <= $data->sheets[0]['numCols']; $j++) {
  }
    //echo "\"".$data->sheets[0]['cells'][$i][4]."\"";

		

        //if($data->sheets[0]['cells'][$i][$j]== $search_var) {
        //if( preg_match("/" . $search_var . "/", $data->sheets[0]['cells'][$i][4]) ) {
         
          //echo "\t<tr>\n";
         // echo "\t\t<td>" . $sheets[0]['numRows'][$i] . "</td><td>" . $sheets[0]['numRows'][$i] . "</td>/n"; 
          // etc...
          //echo "\t</tr>\n";
          //echo $search_var;
        	/*echo "<h3><center>Fail anda berada di .</center></h3> ";
          echo "<br/>";
          echo "\t</tr>\n";*/


         
          

         /*foreach($data->sheets[0]['cells'][$i] as $value)
              {
                   if ($data->sheets[0]['cells'][$i][2])
                    { echo "<center><h3>";
                      echo $data->sheets[0]['cells'][$i][2];
                      echo "</center></h3>";
                    break;
                     echo $data->sheets[0]['cells'][$i][3]; 
                   
                    }
                //echo "\"".$value."\",";
                //echo $value;

              }//foreach*/


        //}
        //echo "\"".$data->sheets[0]['cells'][$i][$j]."\",";

	
}


//print_r($data);
//print_r($data->formatRecords);
?>
