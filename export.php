<?php

//get client data
$result_section = array();
if(trim($query) != '' and Tools::getValue("selectpesquisa") != 'd')
{
	$filter_ = 'c.`'.$filter.'` ';
	//clientes
	$client = new DbQuery();
	$client->select('DISTINCT(c.`id_client`),c.name,c.email,c.phone,c.nif,c.address,c.city,c.postal_code');  
	$client->from('client', 'c');
	$client->join(Shop::addSqlAssociation('client', 'c'));
	$client->where($filter_.' LIKE \'%'.pSQL($query).'%\'');
	
	$client->orderBy('c.`name` ASC');
	$client->groupBy('c.`id_client`');
	$res_client = Db::getInstance()->executeS($client);
	
	$count_resultados = 0;
	if(!empty($res_client)){
		$seccoes[] = array('text' => 'Clientes','val' => 'client');
		if(isset($selectedSections['client']) || $todas_seccoes){
			$count_resultados = count($res_client);
			$result_section['client'] = $res_client;
		}
	}
}   
						
		
//export client data to excel 
$objPHPExcel = new PHPExcel();

$filename = date('m-d-Y_hi').'.xlsx';
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header("Content-Disposition: attachment; filename=\"$filename\" ");
header('Cache-Control: max-age=0');

$objPHPExcel->setActiveSheetIndex(0); 
$objPHPExcel->getActiveSheet()->setTitle('CLIENTES');
$abcd = array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');

$rowCount = 2;
$rowCounth = 0;
$header = false;
if(isset($result_section['client']))
{
	$ids = implode(",", array_column($result_section['client'], 'id_client'));
	$data = array();
	if($ids != '')
	{
		$data = Db::getInstance()->executeS('
			SELECT * FROM `'._DB_PREFIX_.'client` 
			WHERE `id_client` in ('.$ids.')');	
	}
	if(!empty($data))
	{		
		foreach($data as $hils)
		{
			$rowCounth = 0; 
			foreach($hils as $chave => $l){
				if($rowCounth < 26 and !$header)
				{
					$objPHPExcel->getActiveSheet()->SetCellValue($abcd[$rowCounth].'1',  $chave);
				}
						
				if($rowCounth < 26){
					$objPHPExcel->getActiveSheet()->SetCellValue($abcd[$rowCounth].$rowCount, $l);
					$rowCounth++;
				}		
			}
			$header = true;
			$rowCount++;
		}
	}
}

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
