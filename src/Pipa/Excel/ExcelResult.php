<?php

namespace Pipa\Excel;
use PHPExcel;
use PHPExcel_Worksheet;
use Pipa\Dispatch\Result;

class ExcelResultSheet {
	
	private $parent;

	private $sheet;
	
	function __construct(ExcelResult $parent, PHPExcel_Worksheet $sheet) {
		$this->parent = $parent;
		$this->sheet = $sheet;
	}
	
	function cell($coordinate, $value, $type = null) {
		$this->sheet->setCellValue($coordinate, $value);
	}
	
	function tableAt($coordinate, array $records) {
		$row = 0;
		foreach($records as $record) {
			$column = 0;
			foreach($record as $header=>$value) {
				$this->cell($this->coordinate($coordinate, $column, $row), $value);
				$column++;
			}
			$row++;
		}
	}
	
	function coordinate($coordinate, $column = 0, $row = 0) {
		preg_match('/^[A-Z]+/', $coordinate, $c);
		preg_match('/\d+$/', $coordinate, $r);
		return chr(ord($c[0]) + $column) . ($r[0] + $row);
	}
	
	function done() {
		return $this->parent;
	}
}

class ExcelResult extends Result {
	
	private $book;
	private $filename;
	
	function __construct(PHPExcel $book = null) {
		if (!$book)
			$book = new PHPExcel();
		
		$book->removeSheetByIndex(0);
		
		$this->book = $book;
	}
	
	function getBook() {
		return $this->book;
	}
	
	function sheet($title) {
		if (!($sheet = $this->book->getSheetByName($title))) {
			$sheet = $this->book->createSheet();
			$sheet->setTitle($title);
		}
		return new ExcelResultSheet($this, $sheet);
	}
	
	function filename($filename) {
		$this->options['http-download'] = $filename;
		return $this;
	}
}
