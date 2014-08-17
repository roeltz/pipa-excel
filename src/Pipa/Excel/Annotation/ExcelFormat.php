<?php

namespace Pipa\Excel\Annotation;
use Pipa\Dispatch\Annotation\Option;

class ExcelFormat extends Option {
	const EXCEL2007 = "2007";
	const EXCEL97 = "97";
	
	public $name = "excel-format";
}
