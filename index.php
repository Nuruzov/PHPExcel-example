<!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Test parser</title>
</head>
<body>
	

	<?php

			require_once "Classes/PHPExcel.php";

		$tmpfname = "clothes.xlsx";
		$excelReader = PHPExcel_IOFactory::createReaderForFile($tmpfname);
		$excelObj = $excelReader->load($tmpfname);
		$worksheet = $excelObj->getActiveSheet();
		$lastRow = $worksheet->getHighestRow();

		// $excel_writer->setPreCalculateFormulas(false);
		// $excelObj->setPreCalculateFormulas(true);
		


		$myArr = array();

		for ($row=1; $row <=$lastRow ; $row++) {
			

			$myArr[$worksheet->getCell('B'.$row)->getValue()] = $worksheet->getCell('G'.$row)->getCalculatedValue();

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

	?>

</body>
</html>