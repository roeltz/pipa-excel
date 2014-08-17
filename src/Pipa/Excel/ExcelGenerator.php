<?php

namespace Pipa\Excel;
use PHPExcel;
use PHPExcel_Worksheet;

class ExcelGenerator {
	
	private static $defaultMetadata = array();
	
	static function getDefaultMetadata() {
		return self::$defaultMetadata;
	}
	
	static function setDefaultMetadata(array $metadata) {
		self::$defaultMetadata = $metadata;
	}
	
	function generate(array $metadata, array $data) {
		$book = new PHPExcel();
		$this->setMetadata($book, $metadata);
		$this->setData($book, $data);
		return $book;
	}
	
	function setMetadata(PHPExcel $book, array $metadata) {
		$metadata = array_merge(self::$defaultMetadata, $metadata);
		
		$book->getProperties()
			->setCreator(@$metadata['creator'])
			->setLastModifiedBy(@$metadata['lastModifiedBy'])
			->setCreated(@$metadata['created'])
			->setModified(@$metadata['modified'])
			->setTitle(@$metadata['title'])
			->setDescription(@$metadata['description'])
			->setSubject(@$metadata['subject'])
			->setKeywords(@$metadata['keywords'])
			->setCategory(@$metadata['category'])
			->setCompany(@$metadata['company'])
			->setManager(@$metadata['manager'])
		;
	}
	
	function setData(PHPExcel $book, array $data) {
		foreach($data as $sheetName=>$sheetData) {
			$sheet = new PHPExcel_Worksheet($book, $sheetName);
			$book->addSheet($sheet);
			$this->setWorksheetData($sheet, $sheetData['data'], @$sheetData['schema']);
		}
		$book->removeSheetByIndex(0);
	}
	
	function setWorksheetData(PHPExcel_Worksheet $sheet, array $data, array $schema = null) {
		$cells = $data;
		$headers = array();
		$types = array();
		$formats = array();
		
		if ($schema) {
			foreach($schema as $header=>$spec) {
				$headers[] = $header;
				$types[] = isset($spec['type']) ? $spec['type'] : 's';
				$formats[] = isset($spec['format']) ? $spec['format'] : null;
			}
			array_unshift($cells, $headers);
		}
		
		$sheet->fromArray($cells, null, 'A1', true);

		foreach($formats as $i=>$format) {
			if (!is_null($format)) {
				$column = chr(65 + $i);
				$range = "{$column}2:{$column}".(count($data) + 1);
				$sheet->getStyle($range)->getNumberFormat()->setFormatCode($format);
			}
		}
	}
}
