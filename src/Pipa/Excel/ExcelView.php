<?php

namespace Pipa\Excel;
use PHPExcel;
use PHPExcel_Writer_Excel5;
use PHPExcel_Writer_Excel2007;
use Pipa\Dispatch\View;
use Pipa\Dispatch\Dispatch;
use Pipa\Excel\Annotation\ExcelFormat;
use Pipa\HTTP\Response;

class ExcelView implements View {
	
	function render(Dispatch $dispatch) {
		$filename = @$dispatch->result->options['http-download'];
		$format = @$dispatch->result->options['excel-format'];
		
		if (!$filename)
			$filename = "book";

		if (!$format)
			$format = "2007";
		
		if ($dispatch->result->data instanceof PHPExcel) {
			$book = $dispatch->result->data;
		} else if ($dispatch->result instanceof ExcelResult) {
			$book = $dispatch->result->getBook();
		} else {
			$data = $dispatch->result->data;
			$metadata = @$dispatch->result->options['excel-metadata'];
			$generator = new ExcelGenerator();
			
			if (!$metadata)
				$metadata = array();
			
			$book = $generator->generate($metadata, $data);
		}

		$filename = $this->getProperFilename($filename, $format);
		$dispatch->response->setAsDownload("application/vnd.ms-excel", $filename);
		
		$this->output($book, $format);
	}
	
	function getProperFilename($filename, $format) {
		switch($format) {
			case ExcelFormat::EXCEL97:
				return $filename .= '.xls';
			case ExcelFormat::EXCEL2007:
			default:
				return $filename .= '.xlsx';
		}
	}
	
	function output(PHPExcel $book, $format) {
		switch($format) {
			case ExcelFormat::EXCEL97:
				$writer = new PHPExcel_Writer_Excel5($book);
				break;
			case ExcelFormat::EXCEL2007:
			default:
				$writer = new PHPExcel_Writer_Excel2007($book);
		}
		
		$writer->save("php://output");
	}
}
