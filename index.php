<?php
/**
 * PHPExcel - Excel data import to MySQL database script example
 * ==============================================================================
 * 
 * @version v1.0: PHPExcel_excel_to_mysql_demo.php 2016/03/03
 * @copyright Copyright (c) 2016, http://www.ilovephp.net
 * @author Sagar Deshmukh <sagarsdeshmukh91@gmail.com>
 * @SourceOfPHPExcel https://github.com/PHPOffice/PHPExcel, https://sourceforge.net/projects/phpexcelreader/
 * ==============================================================================
 *
 */
 
require_once 'Classes/PHPExcel/IOFactory.php';

// Mysql database
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "test";

$inputfilename = 'example_file.xlsx';
$exceldata = array();
$output = '';

// Create connection
$conn = mysqli_connect($servername, $username, $password, $dbname);
// Check connection
if (!$conn) {
    die("Connection failed: " . mysqli_connect_error());
}
//  Read your Excel workbook
try
{
   /*  $inputfiletype = PHPExcel_IOFactory::identify($inputfilename);
    $objReader = PHPExcel_IOFactory::createReader($inputfiletype);
	$objPHPExcel = $objReader->load($inputfilename); */

	if(isset($_POST["import"])){
			$extension = end(explode(".", $_FILES["excel"]["name"])); // For getting Extension of selected file
			$allowed_extension = array("xls", "xlsx", "csv"); //allowed extension
			
			if(in_array($extension, $allowed_extension)) //check selected file extension is present in allowed extension array
			{
				$file = $_FILES["excel"]["tmp_name"]; // getting temporary source of excel file
				//Aquí es donde seleccionamos nuestro csv
				$fname = $_FILES['excel']['name'];
				echo 'Cargando nombre del archivo: '.$fname.' ';
				//include("PHPExcel/IOFactory.php"); // Add PHPExcel Library in this code
				var_dump(	$_FILES["excel"]);
				$objPHPExcel = PHPExcel_IOFactory::load($file); // create object of PHPExcel library by using load() method and in load method define path of selected file

				$output .= "<label class='text-success'>Data Inserted</label><br /><table class='table table-bordered'>";
				$i=1;
				foreach ($objPHPExcel->getWorksheetIterator() as $worksheet)
				{
					if($i<=1){
						$highestRow = $worksheet->getHighestRow();
						echo $highestRow."<br>";
					//	var_dump( $worksheet);die;
						for($row=2; $row<=$highestRow; $row++)
						{
							$output .= "<tr>";
							$name = mysqli_real_escape_string($conn, $worksheet->getCellByColumnAndRow(0, $row)->getValue());
							$email = mysqli_real_escape_string($conn, $worksheet->getCellByColumnAndRow(1, $row)->getValue());
							if(!empty($name) && !empty($email)){
								$query = "INSERT INTO tbl_excel(excel_name, excel_email) VALUES ('".$name."', '".$email."')";
								mysqli_query($conn, $query);
								
								$output .= '<td>'.$name.'</td>';
								$output .= '<td>'.$email.'</td>';
								$output .= '</tr>';
							}
						} 
					} 
					$i++;
			   }
			  $output .= '</table>';
			}
			else
			{
				$output = '<label class="text-danger">Invalid File</label>'; //if non excel file then
			}
	}
}catch(Exception $e)
{
	
    die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
}

/*//  Get worksheet dimensions
$sheet = $objPHPExcel->getSheet(0); 
$highestRow = $sheet->getHighestRow(); 
$highestColumn = $sheet->getHighestColumn();

//  Loop through each row of the worksheet in turn
for ($row = 1; $row <= $highestRow; $row++)
{ 
    //  Read a row of data into an array
    $rowData = $sheet->rangeToArray('A' . $row . ':' . $highestColumn . $row, NULL, TRUE, FALSE);
	
    //  Insert row data array into your database of choice here
	$sql = "INSERT INTO users (firstname, lastname, email)
			VALUES ('".$rowData[0][0]."', '".$rowData[0][1]."', '".$rowData[0][2]."')";
	
	if (mysqli_query($conn, $sql)) {
		$exceldata[] = $rowData[0];
	} else {
		echo "Error: " . $sql . "<br>" . mysqli_error($conn);
	}
}

// Print excel data
echo "<table>";
foreach ($exceldata as $index => $excelraw)
{
	echo "<tr>";
	foreach ($excelraw as $excelcolumn)
	{
		echo "<td>".$excelcolumn."</td>";
	}
	echo "</tr>";
}
echo "</table>";*/

mysqli_close($conn);
?>

<html>
 <head>
  <title>Import Excel to Mysql using PHPExcel in PHP</title>
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js"></script>
  <link href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" rel="stylesheet" />
  <style>
  body
  {
   margin:0;
   padding:0;
   background-color:#f1f1f1;
  }
  .box
  {
   width:700px;
   border:1px solid #ccc;
   background-color:#fff;
   border-radius:5px;
   margin-top:100px;
  }
  
  </style>
 </head>
 <body>
  <div class="container box">
   <h3 align="center">Import Excel to Mysql using PHPExcel in PHP</h3><br />
   <form method="post" enctype="multipart/form-data">
    <label>Select Excel File</label>
    <input type="file" name="excel" id="excel" />
    <br />
    <input type="submit" name="import" class="btn btn-info" value="Import" />
   </form>
   <br />
   <br />
   <?php
   echo $output;
   ?>
  </div>
 </body>
</html>