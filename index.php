<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Test parser</title>
</head>
<body>
	
<!-- oc_product и oc_product_description -->
	<?php
	error_reporting(E_ALL);
			require_once "Classes/PHPExcel.php";
			require_once "db.php";

		$tmpfname = "test.xlsx";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getActiveSheet();
		$lastRow = $worksheet->getHighestRow();

		// $excel_writer->setPreCalculateFormulas(false);
		// $excelObj->setPreCalculateFormulas(true);
		


		$myArr = array();
		$nameArr = array();

		for ($row=1; $row <=$lastRow ; $row++) {
			

			$myArr[$worksheet->getCell('C'.$row)->getValue()] = $worksheet->getCell('B'.$row)->getCalculatedValue();
			$nameArr[$worksheet->getCell('A'.$row)->getValue()] = $worksheet->getCell('B'.$row)->getValue();

			// getCell('B'.$row) указываем столбец и номер строки
			//getCalculatedValue(); Возвращает высчитанную по формуле данную
			//getValue() Возвращает значение ячейки. Но если там формула к примеру "=F11+E11" то он вернет формулу а не значение
			//getOldCalculatedValue(); Возвращает только те значения где формула а которые сами прописына не возвращает
		}

		

		// echo "<table border='1' cellspacing='2' cellpadding='2'> ";

		// $i=1;
		// foreach ($myArr as $key => $value) {
		// 	echo "<tr>";
		// 	echo "<td>$i</td>";
		// 	echo "<td>$key</td>";
		// 	echo "<td>$value</td>";
		// 	echo "</tr>";
		// 	$i++;
		// }
		// echo "</table>";




		$sql = "
			SELECT oc_product.sku,oc_product.product_id, oc_product.price, oc_product.wholesale_price, oc_product_description.name
			FROM oc_product
			LEFT JOIN oc_product_description ON oc_product.product_id = oc_product_description.product_id
		";

		
		$result = $con->query($sql);
		$ok=0;
		$flag=0;
		$flag2=0;
		$noChange = array();
		$one = 1;

		$stmt = $con->prepare("UPDATE oc_product SET wholesale_price = ?, updated = ? WHERE sku = ?");
		$stmt2 = $con->prepare("UPDATE oc_product SET status = ?, updated = ? WHERE product_id = ?");

		if ($result->num_rows > 0) {
		    // output data of each row
		    while($row = $result->fetch_assoc()) {
		        
		        foreach ($myArr as $key => $value) {
		        	
		        	if($row['sku'] !='' && $key == $row['sku']){
		  				#$upd= "UPDATE oc_product SET wholesale_price = '.$value.' WHERE sku='{$row['sku']}'";
		  				$stmt->bind_param('iss', $value, $one, $row['sku']);
		  				$stmt->execute();
        				
		  				
		        		$flag = 1;

		        		
		        		break;
		        	}else{
		        		$flag = 0;
		        	}
		        }#Конец форич

		        
		        if($flag==0){

		        	foreach ($nameArr as $key => $value) {

		        		if($row['name'] !='' && $key == $row['name']){
		        			$stmt->bind_param('iis', $value, $one, $row['name']);
		  					$stmt->execute();

		  					$flag2 = 1;
		  					
			        		break;
		        		}else{
		        			$flag2 = 0;
		        		}
		        	}#Конец форич

		        	if($flag2 == 0){

		        		$noChange[$ok] = $row['name'];
		        		$ok++;
		        	}else{
		        		$flag2 = 0;
		        	}
		        }#Конец If


		    }
		}
		else {
		    echo "0 results";
		}


		fclose($fp2);
		$stmt->close();
		$stmt2->close();


		$fp = fopen("C:\Users\malik\Desktop\scroll-to-top\counter.txt", "a"); // Открываем файл в режиме записи 
		foreach ($noChange as $key => $value) {
			$text = $value."\r\n";
			$test = fwrite($fp, $text); // Запись в файл
		}


		
		
		fclose($fp); //Закрытие файла

		var_dump($noChange);
	?>

</body>
</html>