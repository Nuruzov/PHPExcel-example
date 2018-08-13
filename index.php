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

		$tmpfname = "clothes.xlsx";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getActiveSheet();
		$lastRow = $worksheet->getHighestRow();

		// $excel_writer->setPreCalculateFormulas(false);
		// $excelObj->setPreCalculateFormulas(true);
		


		$myArr = array();

		for ($row=1; $row <=$lastRow ; $row++) {
			

			$myArr[$worksheet->getCell('L'.$row)->getValue()] = $worksheet->getCell('G'.$row)->getCalculatedValue();

			// getCell('B'.$row) указываем столбец и номер строки
			//getCalculatedValue(); Возвращает высчитанную по формуле данную
			//getValue() Возвращает значение ячейки. Но если там формула к примеру "=F11+E11" то он вернет формулу а не значение
			//getOldCalculatedValue(); Возвращает только те значения где формула а которые сами прописына не возвращает
		}

		

		echo "<table border='1' cellspacing='2' cellpadding='2'> ";

		$i=1;
		foreach ($myArr as $key => $value) {
			echo "<tr>";
			echo "<td>$i</td>";
			echo "<td>$key</td>";
			echo "<td>$value</td>";
			echo "</tr>";
			$i++;
		}
		echo "</table>";




		$sql = "
			SELECT oc_product.sku, oc_product.price, oc_product.wholesale_price, oc_product_description.name
			FROM oc_product
			LEFT JOIN oc_product_description ON oc_product.product_id = oc_product_description.product_id
		";

		
		$result = $con->query($sql);
		$ok=0;
		$flag=0;
		$noChange = array();

		$stmt = $con->prepare("UPDATE oc_product SET wholesale_price = ? WHERE sku = ?");

		if ($result->num_rows > 0) {
		    // output data of each row
		    while($row = $result->fetch_assoc()) {
		        
		        foreach ($myArr as $key => $value) {
		        	
		        	if($row['sku'] !='' && $key == $row['sku']){
		  				#$upd= "UPDATE oc_product SET wholesale_price = '.$value.' WHERE sku='{$row['sku']}'";
		  				$stmt->bind_param('is', $value, $row['sku']);
		  				$stmt->execute();
        				
		  				
		        		$flag==1;

		        		$fp2 = fopen("C:\Users\malik\Desktop\scroll-to-top\counter2.txt", "a");
		        		$text = $row['name']."\r\n";
		        		$test = fwrite($fp2, $text);
		        		fclose($fp2);
		        		break;
		        	}else{
		        		$flag=0;
		        	}
		        }

		        if($flag==0){
		        	$noChange[$ok] = $row['name'];
		        	$ok++;
		        }
		        else{
		        	$flag=0;
		        }
		    }
		} else {
		    echo "0 results";
		}
		$stmt->close();


		$fp = fopen("C:\Users\malik\Desktop\scroll-to-top\counter.txt", "a"); // Открываем файл в режиме записи 
		foreach ($noChange as $key => $value) {
			$text = $value."\r\n";
			$test = fwrite($fp, $text); // Запись в файл
		}
		
		
		fclose($fp); //Закрытие файла


	?>

</body>
</html>