<?php

namespace Pipa\Excel;
use PHPExcel;
use PHPExcel_Worksheet;
use Pipa\Dispatch\Result;

class ExcelResultSheet {
	
	private $parent;
	private $sheet;
	private $offset;
	
	function __construct(ExcelResult $parent, PHPExcel_Worksheet $sheet) {
		$this->parent = $parent;
		$this->sheet = $sheet;
	}
	
	function cell($coordinate, $value, $type = null) {
		if ($this->offset)
			$coordinate = $this->offsetCoordinate($coordinate);
		
		$this->sheet->setCellValue($coordinate, $value);
		return $this;
	}
	
	function merge($range, $value, $type = null) {
		if ($this->offset)
			$range = $this->offsetRange($range);

		$coord = explode(":", $range)[0];
		$this->sheet->mergeCells($range);
		$this->cell($coord, $value);
		return $this;
	}
	
	function tableAt($coordinate, array $records) {
		if ($this->offset)
			$coordinate = $this->offsetCoordinate($coordinate);

		$row = 0;
		foreach($records as $record) {
			$column = 0;
			foreach($record as $header=>$value) {
				$this->cell($this->coordinate($coordinate, $column, $row), $value);
				$column++;
			}
			$row++;
		}
		return $this;
	}
	
	function offset($column, $row) {
		$this->offset = array($column, $row);
		return $this;
	}
	
	function coordinate($coordinate, $column, $row) {
		preg_match('/^[A-Z]+/', $coordinate, $c);
		preg_match('/\d+$/', $coordinate, $r);
		return chr(ord($c[0]) + $column) . ($r[0] + $row);
	}
	
	function offsetCoordinate($coordinate) {
		return $this->coordinate($coordinate, $this->offset[0], $this->offset[1]);
	}
	
	function range($range, $column, $row) {
		list($from, $to) = explode(":", $range);
		return $this->coordinate($from, $column, $row) . ':' . $this->coordinate($to, $column, $row);
	}

	function offsetRange($range) {
		return $this->range($range, $this->offset[0], $this->offset[1]);
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
