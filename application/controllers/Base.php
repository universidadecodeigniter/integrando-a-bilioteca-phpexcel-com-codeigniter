<?php
defined('BASEPATH') OR exit('No direct script access allowed');

class Base extends CI_Controller {

	function __construct(){
		parent::__construct();
		$this->load->library('PHPExcel');
	}

	public function Index()
	{
		// Definindo o nome do arquivo (repare no uso da extensão .php, ela será substituída posteriormente por .xls ou .xlsx)
		$fileName = "PHPExcelFile.php";

		// Definindo o path de salvamento do arquivo
		$saveFilePATH = "./files/".$fileName;

		// Cria um novo objeto
		$objPHPExcel = $this->phpexcel;

		// Define as propriedades do documento
		$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
																 ->setLastModifiedBy("Maarten Balliauw")
																 ->setTitle("PHPExcel Test Document")
																 ->setSubject("PHPExcel Test Document")
																 ->setDescription("Test document for PHPExcel, generated using PHP classes.")
																 ->setKeywords("office PHPExcel php")
																 ->setCategory("Test result file");

		// Insere conteúdo no arquivo
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A1', 'Hello')
																				->setCellValue('B2', 'world!')
																				->setCellValue('C1', 'Hello')
																				->setCellValue('D2', 'world!');

		// Miscellaneous glyphs, UTF-8
		$objPHPExcel->setActiveSheetIndex(0)->setCellValue('A4', 'Miscellaneous glyphs')
																				->setCellValue('A5', 'éàèùâêîôûëïüÿäöüç');

		// Renomeia a worksheet
		$objPHPExcel->getActiveSheet()->setTitle('Simple');

		// Define qual a worksheet estará ativa ao abrir o arquivo
		$objPHPExcel->setActiveSheetIndex(0);

		// Salva o arquivo no formato do Excel 2007 (.xlsx)
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save(str_replace('.php', '.xlsx', $saveFilePATH));

		// Salva outro arquivo no formato do Excel5 (.xls)
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		$objWriter->save(str_replace('.php', '.xls', $saveFilePATH));

		echo date('H:i:s') . " Construção do arquivo concluída" . "<br/>";
		echo 'Arquivo criado em  ' . str_replace('.php', '.xls', $saveFilePATH) . "<br />";
	}

	public function Formulas()
	{
		// Definindo o nome do arquivo (repare no uso da extensão .php, ela será substituída posteriormente por .xls ou .xlsx)
		$fileName = "PHPExcelFormulas.php";

		// Definindo o path de salvamento do arquivo
		$saveFilePATH = "./files/".$fileName;

		// Cria um novo objeto PHPExcel
		$objPHPExcel = $this->phpexcel;

		// Define as propriedades do documento
		$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
																 ->setLastModifiedBy("Maarten Balliauw")
																 ->setTitle("Office 2007 XLSX Test Document")
																 ->setSubject("Office 2007 XLSX Test Document")
																 ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
																 ->setKeywords("office 2007 openxml php")
																 ->setCategory("Test result file");

		// Insere conteúdo no arquivo
		$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Sum:')
																	->setCellValue('B1', 'Range #1')
																	->setCellValue('B2', 3)
																	->setCellValue('B3', 7)
																	->setCellValue('B4', 13)
																	->setCellValue('B5', '=SUM(B2:B4)')
																	->setCellValue('C1', 'Range #2')
																	->setCellValue('C2', 5)
																	->setCellValue('C3', 11)
																	->setCellValue('C4', 17)
																	->setCellValue('C5', '=SUM(C2:C4)')
																	->setCellValue('A7', 'Total of both ranges:')
																	->setCellValue('B7', '=SUM(B5:C5)')
																	->setCellValue('A8', 'Minimum of both ranges:')
																	->setCellValue('B8', '=MIN(B2:C4)')
																	->setCellValue('A9', 'Maximum of both ranges:')
																	->setCellValue('B9', '=MAX(B2:C4)')
																	->setCellValue('A10', 'Average of both ranges:')
																	->setCellValue('B10', '=AVERAGE(B2:C4)');

		// Renomeia a worksheet
		$objPHPExcel->getActiveSheet()->setTitle('Formulas');

		// Define qual a worksheet estará ativa ao abrir o arquivo
		$objPHPExcel->setActiveSheetIndex(0);

		// Salva no formato do Excel 2007 (xlsx)
		$callStartTime = microtime(true);
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
		$objWriter->save(str_replace('.php', '.xlsx', $saveFilePATH));

		// SAlva no formato Excel 95 (xls)
		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		$objWriter->save(str_replace('.php', '.xls', $saveFilePATH));

		echo date('H:i:s') . " Construção do arquivo concluída" . "<br/>";
		echo 'Arquivo criado em  ' . str_replace('.php', '.xls', $saveFilePATH) . "<br />";
	}

}
