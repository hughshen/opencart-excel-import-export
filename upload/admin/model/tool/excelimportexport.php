<?php
static $registry = NULL;

class ModelToolExcelImportExport extends Model {

	private $error = array();

	// import --- product_number => product_id
	private $productsIdMapping = array();
	private $errorProductsIdMapping = array();

	// export --- product_id => product_number
	private $exportProductsIdMapping = array();

	// Sheet title info
	// public $sheetTitle = array("Products","Links","Description","Attribute","Recurring","Discount","Special","Image","Reward","Design","Options");

	// Heading info
	public $productsHeading = array("product_number","model","sku","upc","ean","jan","isbn","mpn","location","quantity","stock_status_id","image","manufacturer_id","shipping","price","points","tax_class_id","date_available","weight","weight_class_id","length","width","height","length_class_id","subtract","minimum","sort_order","status");

	public $linksHeading = array("product_number","category_id","filter_id","store_id","download_id","related_id");

	public $descriptionsHeading = array("product_number","language_id","name","description","tag","meta_title","meta_description","meta_keyword");

	public $attributesHeading = array("product_number", "attribute_id", "language_id", "text");

	public $recurringHeading = array("product_number", "recurring_id", "customer_group_id");

	public $discountsHeading = array("product_number", "customer_group_id", "quantity", "priority", "price", "date_start", "date_end");

	public $specialsHeading = array("product_number", "customer_group_id", "priority", "price", "date_start", "date_end");

	public $imagesHeading = array("product_number", "image", "sort_order");

	public $rewardsHeading = array("product_number", "customer_group_id", "points");

	public $designHeading = array("product_number", "store_id", "layout_id");

	public $optionsHeading = array("product_number","option_id","value","required","option_value_id","quantity","subtract","price","price_prefix","points","points_prefix","weight","weight_prefix");

	function clean( &$str, $allowBlanks=FALSE ) {
		$result = "";
		$n = strlen( $str );
		for ($m=0; $m<$n; $m++) {
			$ch = substr( $str, $m, 1 );
			if (($ch==" ") && (!$allowBlanks) || ($ch=="\n") || ($ch=="\r") || ($ch=="\t") || ($ch=="\0") || ($ch=="\x0B")) {
				continue;
			}
			$result .= $ch;
		}
		return $result;
	}

	function import($filename) {

		// we use our own error handler
		global $registry;
		$registry = $this->registry;
		set_error_handler('error_handler_for_export',E_ALL & !E_NOTICE);
		register_shutdown_function('fatal_error_shutdown_handler_for_export');

		try {
			$database =& $this->db;
			$this->session->data['export_nochange'] = 1;

			// we use the PHPExcel package from http://phpexcel.codeplex.com/
			$cwd = getcwd();
			chdir(DIR_SYSTEM.'PHPExcel');
			require_once('Classes/PHPExcel.php');
			chdir($cwd);

			// parse uploaded spreadsheet file
			$inputFileType = PHPExcel_IOFactory::identify($filename);
			$objReader = PHPExcel_IOFactory::createReader($inputFileType);
			$objReader->setReadDataOnly(true);
			$reader = $objReader->load($filename);

			// read the various worksheets and load them to the database
			$ok = $this->validateImport($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$this->clearCache();
			$this->session->data['export_nochange'] = 0;

			$ok = $this->importProducts( $reader, $database );
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importLinks($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importAttributes($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importRecurring($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importDiscounts($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importSpecials($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importImages($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importRewards($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importDesign($reader, $database);
			if (!$ok) {
				return FALSE;
			}

			$ok = $this->importOptions($reader, $database);
			if (!$ok) {
				return FALSE;
			}
			
			return $ok;
		} catch (Exception $e) {
			$errstr = $e->getMessage();
			$errline = $e->getLine();
			$errfile = $e->getFile();
			$errno = $e->getCode();
			$this->session->data['export_error'] = array( 'errstr'=>$errstr, 'errno'=>$errno, 'errfile'=>$errfile, 'errline'=>$errline );
			if ($this->config->get('config_error_log')) {
				$this->log->write('PHP ' . get_class($e) . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
			}
			return FALSE;
		}
	}

	function exportSample() {

		// we use our own error handler
		global $registry;
		$registry = $this->registry;
		set_error_handler('error_handler_for_export',E_ALL & !E_NOTICE);
		register_shutdown_function('fatal_error_shutdown_handler_for_export');

		// PHPExcel package from http://phpexcel.codeplex.com/
		$cwd = getcwd();
		chdir(DIR_SYSTEM.'PHPExcel');
		require_once('Classes/PHPExcel.php');
		chdir($cwd);

		try {
			// set appropriate timeout limit
			set_time_limit(1800);
			ini_set('memory_limit', '512M');

			$database =& $this->db;

			// create a new workbook
			$workbook = new PHPExcel();

			// set default font name and size
			$workbook->getDefaultStyle()->getFont()->setName('Arial');
			$workbook->getDefaultStyle()->getFont()->setSize(11);
			$workbook->getDefaultStyle()->getAlignment()->setIndent(1);


			// Createing the products sheet
			$workbook->setActiveSheetIndex(0);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Products');
			$this->exportProductsWorksheet($worksheet, $database);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			if (count($this->exportProductsIdMapping) > 0) {

				// creating the Links worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(1);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Links');
				$this->exportLinksWorksheet($worksheet, $database);
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// creating the Descriptions worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(2);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Descriptions');
				$this->exportWorksheet($worksheet, $database, $this->descriptionsHeading, 'product_description');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// creating the Attribute worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(3);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Attribute');
				$this->exportWorksheet($worksheet, $database, $this->attributesHeading, 'product_attribute');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// // creating the Recurring worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(4);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Recurring');
				$this->exportWorksheet($worksheet, $database, $this->recurringHeading, 'product_recurring');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// // creating the Discount worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(5);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Discount');
				$this->exportWorksheet($worksheet, $database, $this->discountsHeading, 'product_discount');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// // creating the Special worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(6);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Special');
				$this->exportWorksheet($worksheet, $database, $this->specialsHeading, 'product_special');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// creating the Image worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(7);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Image');
				$this->exportWorksheet($worksheet, $database, $this->imagesHeading, 'product_image');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// creating the Reward worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(8);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Reward');
				$this->exportWorksheet($worksheet, $database, $this->rewardsHeading, 'product_reward');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// creating the Design worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(9);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Design');
				$this->exportWorksheet($worksheet, $database, $this->designHeading, 'product_to_layout');
				$worksheet->freezePaneByColumnAndRow(1, 2);

				// creating the Options worksheet
				$workbook->createSheet();
				$workbook->setActiveSheetIndex(10);
				$worksheet = $workbook->getActiveSheet();
				$worksheet->setTitle('Options');
				$this->exportOptionsWorksheet($worksheet, $database);
				$worksheet->freezePaneByColumnAndRow(1, 2);
			}

			$workbook->setActiveSheetIndex(0);

			//excelproductsimport execl
      		$datetime = date('Y:m:d');
			header('Content-Type: application/vnd.ms-excel');
			header('Content-Disposition: attachment;filename="Products_Import_Sample_'.$datetime.'.xls"');
			header('Cache-Control: max-age=0');
			$objWriter = PHPExcel_IOFactory::createWriter($workbook, 'Excel5');
			
			$objWriter->save('php://output');

			// Clear the spreadsheet caches
			$this->clearSpreadsheetCache();
			exit;
		} catch (Exception $e) {
			$errstr = $e->getMessage();
			$errline = $e->getLine();
			$errfile = $e->getFile();
			$errno = $e->getCode();
			$this->session->data['export_error'] = array( 'errstr'=>$errstr, 'errno'=>$errno, 'errfile'=>$errfile, 'errline'=>$errline );
			if ($this->config->get('config_error_log')) {
				$this->log->write('PHP ' . get_class($e) . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
			}
			return;
		}
	}

	function exportIdField() {

		// we use our own error handler
		global $registry;
		$registry = $this->registry;
		set_error_handler('error_handler_for_export',E_ALL & !E_NOTICE);
		register_shutdown_function('fatal_error_shutdown_handler_for_export');

		// PHPExcel package from http://phpexcel.codeplex.com/
		$cwd = getcwd();
		chdir(DIR_SYSTEM.'PHPExcel');
		require_once('Classes/PHPExcel.php');
		chdir($cwd);

		try {
			// set appropriate timeout limit
			set_time_limit(1800);
			ini_set('memory_limit', '512M');

			$database =& $this->db;

			// create a new workbook
			$workbook = new PHPExcel();

			// set default font name and size
			$workbook->getDefaultStyle()->getFont()->setName('Arial');
			$workbook->getDefaultStyle()->getFont()->setSize(11);
			$workbook->getDefaultStyle()->getAlignment()->setIndent(1);

			// Get current language id
			$currentLanguageId = (int)$this->config->get("config_language_id");

			// Createing the products sheet
			$workbook->setActiveSheetIndex(0);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Products');
			$this->exportProductsIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Links worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(1);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Links');
			$this->exportLinksIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Descriptions worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(2);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Descriptions');
			$this->exportDescriptionsIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Attribute worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(3);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Attribute');
			$this->exportAttributeIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// // creating the Recurring worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(4);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Recurring');
			$this->exportRecurringIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// // creating the Discount worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(5);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Discount');
			$this->exportDiscountIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// // creating the Special worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(6);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Special');
			$this->exportDiscountIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Image worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(7);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Image');
			// $this->exportImageIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Reward worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(8);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Reward');
			$this->exportDiscountIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Design worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(9);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Design');
			$this->exportDesignIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			// creating the Options worksheet
			$workbook->createSheet();
			$workbook->setActiveSheetIndex(10);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Options');
			$this->exportOptionsIdFieldWorksheet($worksheet, $database, $currentLanguageId);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			$workbook->setActiveSheetIndex(0);

			//excelproductsimport execl
      		$datetime = date('Y:m:d');
			header('Content-Type: application/vnd.ms-excel');
			header('Content-Disposition: attachment;filename="Products_Import_Id_Field'.$datetime.'.xls"');
			header('Cache-Control: max-age=0');
			$objWriter = PHPExcel_IOFactory::createWriter($workbook, 'Excel5');
			
			$objWriter->save('php://output');

			// Clear the spreadsheet caches
			$this->clearSpreadsheetCache();
			exit;
		} catch (Exception $e) {
			$errstr = $e->getMessage();
			$errline = $e->getLine();
			$errfile = $e->getFile();
			$errno = $e->getCode();
			$this->session->data['export_error'] = array( 'errstr'=>$errstr, 'errno'=>$errno, 'errfile'=>$errfile, 'errline'=>$errline );
			if ($this->config->get('config_error_log')) {
				$this->log->write('PHP ' . get_class($e) . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
			}
			return;
		}
	}

	function exportOrdersLocation($data = array()) {

		// 导出订单地理分布会涉及到一些自定义的数据表，比如city,county等等；
		// opencart原本的数据表并不包含以上需要查询的表。

		// we use our own error handler
		global $registry;
		$registry = $this->registry;
		set_error_handler('error_handler_for_export',E_ALL & !E_NOTICE);
		register_shutdown_function('fatal_error_shutdown_handler_for_export');

		// PHPExcel package from http://phpexcel.codeplex.com/
		$cwd = getcwd();
		chdir(DIR_SYSTEM.'PHPExcel');
		require_once('Classes/PHPExcel.php');
		chdir($cwd);

		try {
			// set appropriate timeout limit
			set_time_limit(1800);
			ini_set('memory_limit', '512M');

			$database =& $this->db;

			// create a new workbook
			$workbook = new PHPExcel();

			// set default font name and size
			$workbook->getDefaultStyle()->getFont()->setName('Arial');
			$workbook->getDefaultStyle()->getFont()->setSize(11);
			$workbook->getDefaultStyle()->getAlignment()->setIndent(1);

			// Createing the order location sheet
			$workbook->setActiveSheetIndex(0);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Order Location');
			$this->exportOrdersLocationWorksheet($worksheet, $database, $data);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			$workbook->setActiveSheetIndex(0);

			//excelproductsimport execl
      		$datetime = date('Y:m:d');
			header('Content-Type: application/vnd.ms-excel');
			header('Content-Disposition: attachment;filename="Orders_Location_Export'.$datetime.'.xls"');
			header('Cache-Control: max-age=0');
			$objWriter = PHPExcel_IOFactory::createWriter($workbook, 'Excel5');
			
			$objWriter->save('php://output');

			// Clear the spreadsheet caches
			$this->clearSpreadsheetCache();
			exit;
		} catch (Exception $e) {
			$errstr = $e->getMessage();
			$errline = $e->getLine();
			$errfile = $e->getFile();
			$errno = $e->getCode();
			$this->session->data['export_error'] = array( 'errstr'=>$errstr, 'errno'=>$errno, 'errfile'=>$errfile, 'errline'=>$errline );
			if ($this->config->get('config_error_log')) {
				$this->log->write('PHP ' . get_class($e) . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
			}
			return;
		}
	}

	function exportSalesOrder($data = array()) {

		// 导出订单地理分布会涉及到一些自定义的数据表字段，比如fullname等等；
		// opencart原本的数据表并不包含以上需要查询的表字段。

		// we use our own error handler
		global $registry;
		$registry = $this->registry;
		set_error_handler('error_handler_for_export',E_ALL & !E_NOTICE);
		register_shutdown_function('fatal_error_shutdown_handler_for_export');

		// PHPExcel package from http://phpexcel.codeplex.com/
		$cwd = getcwd();
		chdir(DIR_SYSTEM.'PHPExcel');
		require_once('Classes/PHPExcel.php');
		chdir($cwd);

		try {
			// set appropriate timeout limit
			set_time_limit(1800);
			ini_set('memory_limit', '512M');

			$database =& $this->db;

			// create a new workbook
			$workbook = new PHPExcel();

			// set default font name and size
			$workbook->getDefaultStyle()->getFont()->setName('Arial');
			$workbook->getDefaultStyle()->getFont()->setSize(11);
			$workbook->getDefaultStyle()->getAlignment()->setIndent(1);

			// Createing the order location sheet
			$workbook->setActiveSheetIndex(0);
			$worksheet = $workbook->getActiveSheet();
			$worksheet->setTitle('Sales Order');
			$this->exportSalesOrderWorksheet($worksheet, $database, $data);
			$worksheet->freezePaneByColumnAndRow(1, 2);

			$workbook->setActiveSheetIndex(0);

			//excelproductsimport execl
      		$datetime = date('Y:m:d');
			header('Content-Type: application/vnd.ms-excel');
			header('Content-Disposition: attachment;filename="Sales_Order_Export'.$datetime.'.xls"');
			header('Cache-Control: max-age=0');
			$objWriter = PHPExcel_IOFactory::createWriter($workbook, 'Excel5');
			
			$objWriter->save('php://output');

			// Clear the spreadsheet caches
			$this->clearSpreadsheetCache();
			exit;
		} catch (Exception $e) {
			$errstr = $e->getMessage();
			$errline = $e->getLine();
			$errfile = $e->getFile();
			$errno = $e->getCode();
			$this->session->data['export_error'] = array( 'errstr'=>$errstr, 'errno'=>$errno, 'errfile'=>$errfile, 'errline'=>$errline );
			if ($this->config->get('config_error_log')) {
				$this->log->write('PHP ' . get_class($e) . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
			}
			return;
		}
	}

	function importProducts(&$reader, &$database) {

		$products = $this->getSheetArray($reader, 0, $this->productsHeading);

		$descriptions = $this->getSheetArray($reader, 2, $this->descriptionsHeading);

		// Rearrange descriptions array
		$descriptionsTemp = array();
		foreach ($descriptions as $key => $val) {
			$descriptionsTemp[$val['product_number']][$val['language_id']] = $val;
		}

		$descriptions = $descriptionsTemp;

		// Merge $descriptions and products
		foreach ($products as $key => $val) {
			$products[$key]['product_description'] = $descriptions[$val['product_number']];
		}

		return $this->storeProductsIntoDatabase($database, $products);
	}

	function importLinks(&$reader, &$database) {

		$links = $this->getSheetArray($reader, 1, $this->linksHeading);

		return $this->storeLinksIntoDatabase($database, $links);
	}

	function importAttributes(&$reader, &$database) {

		$attributes = $this->getSheetArray($reader, 3, $this->attributesHeading);

		// Rearrange attributes array
		$attributesTemp = array();
		foreach ($attributes as $key => $val) {
			$attributesTemp[$val['product_number']][$val['language_id']] = $val;
		}

		$attributes = $attributesTemp;

		return $this->storeAttributesIntoDatabase($database, $attributes);
	}

	function importRecurring(&$reader, &$database) {

		$recurring = $this->getSheetArray($reader, 4, $this->recurringHeading);

		return $this->storeRecurringIntoDatabase($database, $recurring);
	}

	function importDiscounts(&$reader, &$database) {

		$discounts = $this->getSheetArray($reader, 5, $this->discountsHeading);

		return $this->storeDiscountsIntoDatabase($database, $discounts);
	}

	function importSpecials(&$reader, &$database) {

		$specials = $this->getSheetArray($reader, 6, $this->specialsHeading);

		return $this->storeSpecialsIntoDatabase($database, $specials);
	}

	function importImages(&$reader, &$database) {

		$images = $this->getSheetArray($reader, 7, $this->imagesHeading);

		return $this->storeImagesIntoDatabase($database, $images);
	}

	function importRewards(&$reader, &$database) {

		$rewards = $this->getSheetArray($reader, 8, $this->rewardsHeading);

		return $this->storeRewardsIntoDatabase($database, $rewards);
	
	}

	function importDesign(&$reader, &$database) {

		$design = $this->getSheetArray($reader, 9, $this->designHeading);

		return $this->storeDesignIntoDatabase($database, $design);
	}

	function importOptions(&$reader, &$database) {

		$options = $this->getSheetArray($reader, 10, $this->optionsHeading);

		// Check the datetime format in value
		foreach ($options as $key => $val) {

			if (preg_match('/^[\d]{4,}$/si', $val['value'])) {
				$options[$key]['value'] = date('Y-m-d', PHPExcel_Shared_Date::ExcelToPHP($val['value']));
				continue;
			}
			if (preg_match('/^0\.[\d]{6,}$/si', $val['value'])) {
				$options[$key]['value'] = date('G:i', PHPExcel_Shared_Date::ExcelToPHP($val['value']));
				continue;
			}
			if (preg_match('/[\d]{4,}+\.[\d]{6,}$/si', $val['value'])) {
				$options[$key]['value'] = date('Y-m-d G:i', PHPExcel_Shared_Date::ExcelToPHP($val['value']));
				continue;
			}
		}

		return $this->storeOptionsIntoDatabase($database, $options);
	}

	function storeProductsIntoDatabase(&$database, &$products) {

		$currentStoreId = $this->config->get('config_store_id');

		foreach ($products as $key => $data) {

			$database->query("SET AUTOCOMMIT=0;");
			$database->query("START TRANSACTION;");

			try {
				$database->query("INSERT INTO " . DB_PREFIX . "product SET model = '" . $database->escape($data['model']) . "', sku = '" . $database->escape($data['sku']) . "', upc = '" . $database->escape($data['upc']) . "', ean = '" . $database->escape($data['ean']) . "', jan = '" . $database->escape($data['jan']) . "', isbn = '" . $database->escape($data['isbn']) . "', mpn = '" . $database->escape($data['mpn']) . "', location = '" . $database->escape($data['location']) . "', quantity = '" . (int)$data['quantity'] . "', minimum = '" . (int)$data['minimum'] . "', subtract = '" . (int)$data['subtract'] . "', stock_status_id = '" . (int)$data['stock_status_id'] . "', date_available = '" . $database->escape($data['date_available']) . "', manufacturer_id = '" . (int)$data['manufacturer_id'] . "', shipping = '" . (int)$data['shipping'] . "', price = '" . (float)$data['price'] . "', points = '" . (int)$data['points'] . "', weight = '" . (float)$data['weight'] . "', weight_class_id = '" . (int)$data['weight_class_id'] . "', length = '" . (float)$data['length'] . "', width = '" . (float)$data['width'] . "', height = '" . (float)$data['height'] . "', length_class_id = '" . (int)$data['length_class_id'] . "', status = '" . (int)$data['status'] . "', tax_class_id = '" . (int)$data['tax_class_id'] . "', sort_order = '" . (int)$data['sort_order'] . "', date_added = NOW(), date_modified = NOW()");

				$product_id = $database->getLastId();

				if (isset($data['image'])) {
					$database->query("UPDATE " . DB_PREFIX . "product SET image = '" . $database->escape($data['image']) . "' WHERE product_id = '" . (int)$product_id . "'");
				}

				$database->query("INSERT INTO " . DB_PREFIX . "product_to_store SET product_id = '" . (int)$product_id . "', store_id = '" . (int)$currentStoreId . "'");

				foreach ($data['product_description'] as $language_id => $v) {
					$database->query("INSERT INTO " . DB_PREFIX . "product_description SET product_id = '" . (int)$product_id . "', language_id = '" . (int)$language_id . "', name = '" . $database->escape($v['name']) . "', description = '" . $database->escape($v['description']) . "', tag = '" . $database->escape($v['tag']) . "', meta_title = '" . $database->escape($v['meta_title']) . "', meta_description = '" . $database->escape($v['meta_description']) . "', meta_keyword = '" . $database->escape($v['meta_keyword']) . "'");
				}

				$database->query("COMMIT;");

				$this->productsIdMapping[$data['product_number']] = $product_id;
			} catch (Exception $e) {
				$database->query("ROOLBACK;");
				$this->errorProductsIdMapping[$data['product_number']];
				error_log(date('Y-m-d H:i:s - ', time()).$e->getMessage()."\n",3,DIR_LOGS."error.txt");
			}
		}

		return TRUE;
	}

	function storeLinksIntoDatabase(&$database, &$links) {

		foreach ($links as $key => $data) {

			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {

				$product_id = $this->productsIdMapping[$data['product_number']];

				if (isset($data['category_id'])) {
					$database->query("INSERT INTO " . DB_PREFIX . "product_to_category SET product_id = '" . (int)$product_id . "', category_id = '" . (int)$data['category_id'] . "'");
				}

				if (isset($data['filter_id'])) {
					$database->query("INSERT INTO " . DB_PREFIX . "product_filter SET product_id = '" . (int)$product_id . "', filter_id = '" . (int)$data['filter_id'] . "'");
				}

				if (isset($data['store_id'])) {
					$database->query("INSERT INTO " . DB_PREFIX . "product_to_store SET product_id = '" . (int)$product_id . "', store_id = '" . (int)$data['store_id'] . "'");
				}

				if (isset($data['download_id'])) {
					$database->query("INSERT INTO " . DB_PREFIX . "product_to_download SET product_id = '" . (int)$product_id . "', download_id = '" . (int)$data['download_id'] . "'");
				}

				if (isset($data['related_id'])) {
					$database->query("DELETE FROM " . DB_PREFIX . "product_related WHERE product_id = '" . (int)$product_id . "' AND related_id = '" . (int)$data['related_id'] . "'");
					$database->query("INSERT INTO " . DB_PREFIX . "product_related SET product_id = '" . (int)$product_id . "', related_id = '" . (int)$data['related_id'] . "'");
					$database->query("DELETE FROM " . DB_PREFIX . "product_related WHERE product_id = '" . (int)$data['related_id'] . "' AND related_id = '" . (int)$product_id . "'");
					$database->query("INSERT INTO " . DB_PREFIX . "product_related SET product_id = '" . (int)$data['related_id'] . "', related_id = '" . (int)$product_id . "'");
				}
			}
		}

		return TRUE;
	}

	function storeAttributesIntoDatabase(&$database, &$attributes) {

		foreach ($attributes as $product_number => $attribute) {
			if (isset($this->productsIdMapping[$product_number]) && (int)$this->productsIdMapping[$product_number] > 0) {

				$product_id = $this->productsIdMapping[$product_number];
				foreach ($attribute as $language_id => $data) {
					$database->query("INSERT INTO " . DB_PREFIX . "product_attribute SET product_id = '" . (int)$product_id . "', attribute_id = '" . (int)$data['attribute_id'] . "', language_id = '" . (int)$language_id . "', text = '" .  $database->escape($data['text']) . "'");
				}
			}
		}

		return TRUE;
	}

	function storeRecurringIntoDatabase(&$database, &$recurring) {

		foreach ($recurring as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];
				$database->query("INSERT INTO `" . DB_PREFIX . "product_recurring` SET `product_id` = " . (int)$product_id . ", customer_group_id = " . (int)$data['customer_group_id'] . ", `recurring_id` = " . (int)$data['recurring_id']);
			}
		}

		return TRUE;
	}

	function storeDiscountsIntoDatabase(&$database, &$discounts) {

		foreach ($discounts as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];
				$database->query("INSERT INTO " . DB_PREFIX . "product_discount SET product_id = '" . (int)$product_id . "', customer_group_id = '" . (int)$data['customer_group_id'] . "', quantity = '" . (int)$data['quantity'] . "', priority = '" . (int)$data['priority'] . "', price = '" . (float)$data['price'] . "', date_start = '" . $database->escape($data['date_start']) . "', date_end = '" . $database->escape($data['date_end']) . "'");
			}
		}
		
		return TRUE;
	}

	function storeSpecialsIntoDatabase(&$database, &$specials) {
		
		foreach ($specials as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];
				$database->query("INSERT INTO " . DB_PREFIX . "product_special SET product_id = '" . (int)$product_id . "', customer_group_id = '" . (int)$data['customer_group_id'] . "', priority = '" . (int)$data['priority'] . "', price = '" . (float)$data['price'] . "', date_start = '" . $database->escape($data['date_start']) . "', date_end = '" . $database->escape($data['date_end']) . "'");
			}
		}

		return TRUE;
	}

	function storeImagesIntoDatabase(&$database, &$images) {

		foreach ($images as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];
				$database->query("INSERT INTO " . DB_PREFIX . "product_image SET product_id = '" . (int)$product_id . "', image = '" . $database->escape($data['image']) . "', sort_order = '" . (int)$data['sort_order'] . "'");
			}
		}

		return TRUE;
	}

	function storeRewardsIntoDatabase(&$database, &$rewards) {

		foreach ($rewards as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];
				$database->query("INSERT INTO " . DB_PREFIX . "product_reward SET product_id = '" . (int)$product_id . "', customer_group_id = '" . (int)$data['customer_group_id'] . "', points = '" . (int)$data['points'] . "'");
			}
		}

		return TRUE;
	}

	function storeDesignIntoDatabase(&$database, &$design) {

		foreach ($design as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];
				$database->query("INSERT INTO " . DB_PREFIX . "product_to_layout SET product_id = '" . (int)$product_id . "', store_id = '" . (int)$data['store_id'] . "', layout_id = '" . (int)$data['layout_id'] . "'");
			}
		}

		return TRUE;
	}

	function storeOptionsIntoDatabase(&$database, &$options) {

		foreach ($options as $key => $data) {
			if (isset($this->productsIdMapping[$data['product_number']]) && (int)$this->productsIdMapping[$data['product_number']] > 0) {
				$product_id = $this->productsIdMapping[$data['product_number']];

				// if option type equal 'select' 'radio' 'checkbox' 'image'
				if ((int)$data['option_id'] == 1 || (int)$data['option_id'] == 2 || (int)$data['option_id'] == 5 || (int)$data['option_id'] == 13) {

					try {

						$database->query("SET AUTOCOMMIT=0;");
						$database->query("START TRANSACTION;");

						$database->query("INSERT INTO " . DB_PREFIX . "product_option SET product_id = '" . (int)$product_id . "', option_id = '" . (int)$data['option_id'] . "', required = '" . (int)$data['required'] . "'");

						$product_option_id = $database->getLastId();

						$database->query("INSERT INTO " . DB_PREFIX . "product_option_value SET product_option_id = '" . (int)$product_option_id . "', product_id = '" . (int)$product_id . "', option_id = '" . (int)$data['option_id'] . "', option_value_id = '" . (int)$data['option_value_id'] . "', quantity = '" . (int)$data['quantity'] . "', subtract = '" . (int)$data['subtract'] . "', price = '" . (float)$data['price'] . "', price_prefix = '" . $database->escape($data['price_prefix']) . "', points = '" . (int)$data['points'] . "', points_prefix = '" . $database->escape($data['points_prefix']) . "', weight = '" . (float)$data['weight'] . "', weight_prefix = '" . $database->escape($data['weight_prefix']) . "'");

						$database->query("COMMIT;");
					} catch (Exception $e) {
						$database->query("ROOLBACK;");
						error_log(date('Y-m-d H:i:s - ', time()).$e->getMessage()."\n",3,DIR_LOGS."error.txt");
					}
				} else {
					$database->query("INSERT INTO " . DB_PREFIX . "product_option SET product_id = '" . (int)$product_id . "', option_id = '" . (int)$data['option_id'] . "', value = '" . $database->escape($data['value']) . "', required = '" . (int)$data['required'] . "'");
				}
			}
		}

		return TRUE;
	}

	function getCell(&$worksheet, $row, $col, $default_val='') {

		$col -= 1; // we use 1-based, PHPExcel uses 0-based column index
		$row += 1; // we use 0-based, PHPExcel used 1-based row index

		return ($worksheet->cellExistsByColumnAndRow($col,$row)) ? $worksheet->getCellByColumnAndRow($col,$row)->getValue() : $default_val;
	}

	function getSheetArray(&$reader, $index, $heading) {

		// Get sheet array
		$index = (int)$index;
		if ($index < 0) $index = 0;
		if ($index >= $reader->getSheetCount()) $index = $reader->getSheetCount() - 1;

		$sheet = $reader->getSheet($index);
		$rows = $sheet->getHighestRow();
		$cols = PHPExcel_Cell::columnIndexFromString($sheet->getHighestColumn());

		$arr = array();

		if (count($heading) == 0) return $arr;

		for ($i = 0; $i < $rows; $i++) {
			if ($i == 0) continue;
			for ($j = 0; $j < $cols; $j++) {
				$arr[$i-1][$heading[$j]] = $this->getCell($sheet, $i, $j+1);

				if (($heading[$j] == 'date_start' || $heading[$j] == 'date_end' || $heading[$j] == 'date_available') && $arr[$i-1][$heading[$j]] != '0000-00-00') {
					$arr[$i-1][$heading[$j]] = date('Y-m-d', PHPExcel_Shared_Date::ExcelToPHP($arr[$i-1][$heading[$j]]));
				}
			}
		}

		return $arr;
	}

	function validateHeading(&$data, &$expected) {

		$heading = array();
		$k = PHPExcel_Cell::columnIndexFromString( $data->getHighestColumn() );
		if ($k != count($expected)) {
			return FALSE;
		}
		$i = 0;
		for ($j=1; $j <= $k; $j+=1) {
			$heading[] = $this->getCell($data,$i,$j);
		}
		$valid = TRUE;
		for ($i=0; $i < count($expected); $i+=1) {
			if (!isset($heading[$i])) {
				$valid = FALSE;
				break;
			}
			if (strtolower($heading[$i]) != strtolower($expected[$i])) {
				$valid = FALSE;
				break;
			}
		}
		return $valid;
	}

	function validateProducts(&$reader) {

		$data =& $reader->getSheet(0);
		return $this->validateHeading($data, $this->productsHeading);
	}


	function validateLinks(&$reader) {
		
		$data =& $reader->getSheet(1);
		return $this->validateHeading($data, $this->linksHeading);
	}

	function validateDescription(&$reader) {
		
		$data =& $reader->getSheet(2);
		return $this->validateHeading($data, $this->descriptionsHeading);
	}

	function validateAttributes(&$reader) {
		
		$data =& $reader->getSheet(3);
		return $this->validateHeading($data, $this->attributesHeading);
	}

	function validateRecurring(&$reader) {
		
		$data =& $reader->getSheet(4);
		return $this->validateHeading($data, $this->recurringHeading);
	}

	function validateDiscounts(&$reader) {
		
		$data =& $reader->getSheet(5);
		return $this->validateHeading($data, $this->discountsHeading);
	}

	function validateSpecials(&$reader) {
		
		$data =& $reader->getSheet(6);
		return $this->validateHeading($data, $this->specialsHeading);
	}

	function validateImages(&$reader) {
		
		$data =& $reader->getSheet(7);
		return $this->validateHeading($data, $this->imagesHeading);
	}

	function validateRewards(&$reader) {
		
		$data =& $reader->getSheet(8);
		return $this->validateHeading($data, $this->rewardsHeading);
	}

	function validateDesign(&$reader) {
		
		$data =& $reader->getSheet(9);
		return $this->validateHeading($data, $this->designHeading);
	}

	function validateOptions(&$reader) {

		$data =& $reader->getSheet(10);
		return $this->validateHeading($data, $this->optionsHeading);
	}

	function validateSomeIdField($reader, $database) {

		// Validate products sheet
		$products = $this->getSheetArray($reader, 0, $this->productsHeading);

		$result = array();
		$tempId = $database->query("SELECT stock_status_id FROM " . DB_PREFIX . "stock_status")->rows;
		foreach ($tempId as $key => $val) {
			$result['stock_status_id'][] = $val['stock_status_id'];
		}
		$result['stock_status_id'] = array_unique($result['stock_status_id']);

		$tempId = $database->query("SELECT manufacturer_id FROM " . DB_PREFIX . "manufacturer")->rows;
		foreach ($tempId as $key => $val) {
			$result['manufacturer_id'][] = $val['manufacturer_id'];
		}

		$tempId = $database->query("SELECT tax_class_id FROM " . DB_PREFIX . "tax_class")->rows;
		foreach ($tempId as $key => $val) {
			$result['tax_class_id'][] = $val['tax_class_id'];
		}

		$tempId = $database->query("SELECT weight_class_id FROM " . DB_PREFIX . "weight_class")->rows;
		foreach ($tempId as $key => $val) {
			$result['weight_class_id'][] = $val['weight_class_id'];
		}

		$tempId = $database->query("SELECT length_class_id FROM " . DB_PREFIX . "length_class")->rows;
		foreach ($tempId as $key => $val) {
			$result['length_class_id'][] = $val['length_class_id'];
		}

		// Start validate products sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($products as $key => $val) {
				if (!in_array($val[$idField], $idArray) && $val[$idField] != 0) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 1 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate links sheet
		$links = $this->getSheetArray($reader, 1, $this->linksHeading);
		$result = array();
		$tempId = $database->query("SELECT category_id FROM " . DB_PREFIX . "category")->rows;
		foreach ($tempId as $key => $val) {
			$result['category_id'][] = $val['category_id'];
		}

		$tempId = $database->query("SELECT filter_id FROM " . DB_PREFIX . "filter")->rows;
		foreach ($tempId as $key => $val) {
			$result['filter_id'][] = $val['filter_id'];
		}

		$tempId = $database->query("SELECT store_id FROM " . DB_PREFIX . "store")->rows;
		foreach ($tempId as $key => $val) {
			$result['store_id'][] = $val['store_id'];
		}

		$tempId = $database->query("SELECT download_id FROM " . DB_PREFIX . "download")->rows;
		foreach ($tempId as $key => $val) {
			$result['download_id'][] = $val['download_id'];
		}

		$tempId = $database->query("SELECT product_id FROM " . DB_PREFIX . "product")->rows;
		foreach ($tempId as $key => $val) {
			$result['related_id'][] = $val['product_id'];
		}

		// Start validate links sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($links as $key => $val) {
				if (!in_array($val[$idField], $idArray) && $val[$idField] != 0) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 2 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate descriptions sheet
		$descriptions = $this->getSheetArray($reader, 2, $this->descriptionsHeading);
		$result = array();
		$tempId = $database->query("SELECT language_id FROM " . DB_PREFIX . "language")->rows;
		foreach ($tempId as $key => $val) {
			$result['language_id'][] = $val['language_id'];
		}

		// Start validate descriptions sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($descriptions as $key => $val) {
				if (!in_array($val[$idField], $idArray)) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 3 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate attributes sheet
		$attributes = $this->getSheetArray($reader, 3, $this->attributesHeading);
		$result = array();
		$tempId = $database->query("SELECT attribute_id FROM " . DB_PREFIX . "attribute")->rows;
		foreach ($tempId as $key => $val) {
			$result['attribute_id'][] = $val['attribute_id'];
		}

		$tempId = $database->query("SELECT language_id FROM " . DB_PREFIX . "language")->rows;
		foreach ($tempId as $key => $val) {
			$result['language_id'][] = $val['language_id'];
		}

		// Start validate attributes sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($attributes as $key => $val) {
				if (!in_array($val[$idField], $idArray)) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 4 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate recurring sheet
		$recurring = $this->getSheetArray($reader, 4, $this->recurringHeading);
		$result = array();
		$tempId = $database->query("SELECT recurring_id FROM " . DB_PREFIX . "recurring")->rows;
		foreach ($tempId as $key => $val) {
			$result['recurring_id'][] = $val['recurring_id'];
		}

		$tempId = $database->query("SELECT customer_group_id FROM " . DB_PREFIX . "customer_group")->rows;
		foreach ($tempId as $key => $val) {
			$result['customer_group_id'][] = $val['customer_group_id'];
		}

		// Start validate recurring sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($recurring as $key => $val) {
				if (!in_array($val[$idField], $idArray)) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 5 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate discounts sheet
		$discounts = $this->getSheetArray($reader, 5, $this->discountsHeading);
		$result = array();
		$tempId = $database->query("SELECT customer_group_id FROM " . DB_PREFIX . "customer_group")->rows;
		foreach ($tempId as $key => $val) {
			$result['customer_group_id'][] = $val['customer_group_id'];
		}

		// Start validate discounts sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($discounts as $key => $val) {
				if (!in_array($val[$idField], $idArray)) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 6 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate specials sheet
		$specials = $this->getSheetArray($reader, 6, $this->specialsHeading);
		$result = array();
		$tempId = $database->query("SELECT customer_group_id FROM " . DB_PREFIX . "customer_group")->rows;
		foreach ($tempId as $key => $val) {
			$result['customer_group_id'][] = $val['customer_group_id'];
		}

		// Start validate specials sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($specials as $key => $val) {
				if (!in_array($val[$idField], $idArray)) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 7 $idField"."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate rewards sheet
		$rewards = $this->getSheetArray($reader, 8, $this->rewardsHeading);
		$result = array();
		$tempId = $database->query("SELECT customer_group_id FROM " . DB_PREFIX . "customer_group")->rows;
		foreach ($tempId as $key => $val) {
			$result['customer_group_id'][] = $val['customer_group_id'];
		}

		// Start validate rewards sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($rewards as $key => $val) {
				if (!in_array($val[$idField], $idArray)) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 9 $idField"."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate design sheet
		$design = $this->getSheetArray($reader, 9, $this->designHeading);
		$result = array();
		$tempId = $database->query("SELECT store_id FROM " . DB_PREFIX . "store")->rows;
		foreach ($tempId as $key => $val) {
			$result['store_id'][] = $val['store_id'];
		}

		$tempId = $database->query("SELECT layout_id FROM " . DB_PREFIX . "layout")->rows;
		foreach ($tempId as $key => $val) {
			$result['layout_id'][] = $val['layout_id'];
		}

		// Start validate design sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($design as $key => $val) {
				if (!in_array($val[$idField], $idArray) && $val[$idField] != 0) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 10 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		// Validate options sheet
		$options = $this->getSheetArray($reader, 10, $this->optionsHeading);
		$result = array();
		$tempId = $database->query("SELECT option_id FROM " . DB_PREFIX . "option")->rows;
		foreach ($tempId as $key => $val) {
			$result['option_id'][] = $val['option_id'];
		}

		$tempId = $database->query("SELECT option_value_id FROM " . DB_PREFIX . "option_value")->rows;
		foreach ($tempId as $key => $val) {
			$result['option_value_id'][] = $val['option_value_id'];
		}

		// Start validate options sheet id
		foreach ($result as $idField => $idArray) {
			foreach ($options as $key => $val) {
				if (!in_array($val[$idField], $idArray) && $val[$idField] != 0) {
					error_log(date('Y-m-d H:i:s - ', time())."Something wrong in sheet 11 where product_number=".++$key." ".$idField."\n",3,DIR_LOGS."error.txt");
					return FALSE;
				}
			}
		}

		return TRUE;
	}

	function validateImport(&$reader, $database)
	{
		if ($reader->getSheetCount() != 11) {
			error_log(date('Y-m-d H:i:s - ', time())."The sheets count should be 11.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateProducts($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The products sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateLinks($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The links sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateDescription($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The descriptions sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateAttributes($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The attributes sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateRecurring($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The recurring sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateDiscounts($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The discounts sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateSpecials($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The specials sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateImages($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The images sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateRewards($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The rewards sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateDesign($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The design sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateOptions($reader)) {
			error_log(date('Y-m-d H:i:s - ', time())."The options sheet title is incorrect.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}
		if (!$this->validateSomeIdField($reader, $database)) {
			error_log(date('Y-m-d H:i:s - ', time())."Some id field is not available.\n",3,DIR_LOGS."error.txt");
			return FALSE;
		}

		return TRUE;
	}

	protected function setCell(&$worksheet, $row/*1-based*/, $col/*0-based*/, $val, $style=NULL) {

		$worksheet->setCellValueByColumnAndRow( $col, $row, $val );
		if ($style) {
			$worksheet->getStyleByColumnAndRow($col,$row)->applyFromArray($style);
		}
	}

	function exportProductsWorksheet(&$worksheet, &$database) {
		
		// Set the column widths
		foreach ($this->productsHeading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);
		
		$products = $database->query("SELECT * FROM ". DB_PREFIX . "product LIMIT 20")->rows;

		// export data info to sheet
		foreach ($products as $key => $val) {
			$worksheet->getRowDimension($key+2)->setRowHeight(26);
			$i = 0;
			foreach ($this->productsHeading as $k => $v) {
				if ($v == 'product_number') {
					$this->setCell($worksheet, $key+2, $i, $key+1);
				} else {
					$this->setCell($worksheet, $key+2, $i, $products[$key][$v]);
				}
				$i++;
			}
			$this->exportProductsIdMapping[$val['product_id']] = $key + 1;
		}
	}

	function exportLinksWorksheet(&$worksheet, &$database) {

		// Set the column widths
		foreach ($this->linksHeading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();
		foreach ($this->exportProductsIdMapping as $product_id => $product_number) {

			$data[$product_id]['product_number'] = $product_number;

			$tmp = $database->query("SELECT category_id FROM ". DB_PREFIX . "product_to_category WHERE product_id = " . $product_id)->row;
			if (isset($tmp['category_id'])) {
				$data[$product_id]['category_id'] = $tmp['category_id'];
			} else {
				$data[$product_id]['category_id'] = NULL;
			}

			$tmp = $database->query("SELECT filter_id FROM ". DB_PREFIX . "product_filter WHERE product_id = " . $product_id)->row;
			if (isset($tmp['filter_id'])) {
				$data[$product_id]['filter_id'] = $tmp['filter_id'];
			} else {
				$data[$product_id]['filter_id'] = NULL;
			}

			$tmp = $database->query("SELECT store_id FROM ". DB_PREFIX . "product_to_store WHERE product_id = " . $product_id)->row;
			if (isset($tmp['store_id'])) {
				$data[$product_id]['store_id'] = $tmp['store_id'];
			} else {
				$data[$product_id]['store_id'] = NULL;
			}

			$tmp = $database->query("SELECT download_id FROM ". DB_PREFIX . "product_to_download WHERE product_id = " . $product_id)->row;
			if (isset($tmp['download_id'])) {
				$data[$product_id]['download_id'] = $tmp['download_id'];
			} else {
				$data[$product_id]['download_id'] = NULL;
			}

			$tmp = $database->query("SELECT related_id FROM ". DB_PREFIX . "product_related WHERE product_id = " . $product_id)->row;
			if (isset($tmp['related_id'])) {
				$data[$product_id]['related_id'] = $tmp['related_id'];
			} else {
				$data[$product_id]['related_id'] = NULL;
			}
		}

		// Export data info to sheet
		// Start in row 2
		$key = 2;
		foreach ($data as $val) {
			$worksheet->getRowDimension($key)->setRowHeight(26);
			$i = 0;
			foreach ($val as $k => $v) {
				$this->setcell($worksheet, $key, $i, $v);
				$i++;
			}
			$key++;
		}
	}


	function exportOptionsWorksheet(&$worksheet, &$database) {
		// Set the column widths
		foreach ($this->optionsHeading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);
		// o.option_id o.value o.required ov.option_value_id ov.quantity ov.subtract ov.price ov.price_prefix ov.points ov.points_prefix ov.weight ov.weight_prefix
		$data = array();
		foreach ($this->exportProductsIdMapping as $product_id => $product_number) {
			$data[] = $database->query("SELECT * FROM " . DB_PREFIX . "product_option WHERE option_id NOT IN (1, 2, 5,13) AND product_id = " . $product_id)->rows;

			$data[] = $database->query("SELECT o.product_id, o.option_id, o.value, o.required, ov.option_value_id, ov.quantity, ov.subtract, ov.price, ov.price_prefix, ov.points, ov.points_prefix, ov.weight, ov.weight_prefix FROM " . DB_PREFIX . "product_option o LEFT JOIN " . DB_PREFIX . "product_option_value ov ON (o.option_id = ov.option_id) WHERE o.option_id IN (1, 2, 5,13) AND o.product_id = " . $product_id)->rows;
		}
		// Rearrange
		$dataTmp = array();
		foreach ($data as $key => $val) {
			if (count($val) > 0) {
				foreach ($val as $v) {
					$dataTmp[] = $v;
				}
			}
		}
		$data = $dataTmp;

		// Export data info to sheet
		// Start in row 2
		$key = 2;
		foreach ($data as $d) {
			foreach ($this->optionsHeading as $k => $v) {
				if ($k == 0) {
					$this->setCell($worksheet, $key, $k, $this->exportProductsIdMapping[$d['product_id']]);
				} else {
					// if option type equal 'select' 'radio' 'checkbox' 'image'
					if (array_key_exists($v, $d)) {
						// echo $d[$v].'--';
						$this->setCell($worksheet, $key, $k, $d[$v]);
					} else {
						break;
					}
				}
			}
			$key++;
		}
		// exit;
	}

	function exportWorksheet(&$worksheet, &$database, $heading, $table) {
		// Set the column widths
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();
		foreach ($this->exportProductsIdMapping as $product_id => $product_number) {
			$data[] = $database->query("SELECT * FROM " . DB_PREFIX . $table . " WHERE product_id = " . $product_id)->rows;
		}
		// Rearrange
		$dataTmp = array();
		foreach ($data as $key => $val) {
			if (count($val) > 0) {
				foreach ($val as $v) {
					$dataTmp[] = $v;
				}
			}
		}
		$data = $dataTmp;

		// Export data info to sheet
		// Start in row 2
		$key = 2;
		foreach ($data as $d) {
			foreach ($heading as $k => $v) {
				if ($k == 0) {
					$this->setCell($worksheet, $key, $k, $this->exportProductsIdMapping[$d['product_id']]);
				} else {
					$this->setCell($worksheet, $key, $k, htmlspecialchars_decode($d[$v]));
				}
			}
			$key++;
		}
	}

	function exportProductsIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("stock_status_id", "stock_status_name", "manufacturer_id", "manufacturer_name", "tax_class_id", "tax_class_name", "weight_class_id", "weight_class_name", "length_class_id", "length_class_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();
		$data[] = $database->query("SELECT stock_status_id, name FROM " . DB_PREFIX . "stock_status WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT manufacturer_id, name FROM " . DB_PREFIX . "manufacturer")->rows;

		$data[] = $database->query("SELECT tax_class_id, title FROM " . DB_PREFIX . "tax_class")->rows;

		$data[] = $database->query("SELECT weight_class_id, title FROM " . DB_PREFIX . "weight_class_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT length_class_id, title FROM " . DB_PREFIX . "length_class_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportLinksIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("category_id", "category_name", "filter_id", "filter_name", "store_id", "store_name", "download_id", "download_name", "related_id", "related_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();
		$data[] = $database->query("SELECT category_id, name FROM " . DB_PREFIX . "category_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT filter_id, name FROM " . DB_PREFIX . "filter_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT store_id, name FROM " . DB_PREFIX . "store")->rows;

		$data[] = $database->query("SELECT download_id, name FROM " . DB_PREFIX . "download_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT product_id, name FROM " . DB_PREFIX . "product_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportDescriptionsIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("language_id", "language_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();
		$data[] = $database->query("SELECT language_id, name FROM " . DB_PREFIX . "language")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportAttributeIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("attribute_id", "attribute_name", "language_id", "language_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();

		$data[] = $database->query("SELECT attribute_id, name FROM " . DB_PREFIX . "attribute_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT language_id, name FROM " . DB_PREFIX . "language")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportRecurringIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("recurring_id", "recurring_name", "customer_group_id", "customer_group_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();

		$data[] = $database->query("SELECT recurring_id, name FROM " . DB_PREFIX . "recurring_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT customer_group_id, name FROM " . DB_PREFIX . "customer_group_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportDiscountIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("customer_group_id", "customer_group_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();

		$data[] = $database->query("SELECT customer_group_id, name FROM " . DB_PREFIX . "customer_group_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportDesignIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("store_id", "store_name", "layout_id", "layout_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();

		$data[] = $database->query("SELECT store_id, name FROM " . DB_PREFIX . "store")->rows;

		$data[] = $database->query("SELECT layout_id, name FROM " . DB_PREFIX . "layout")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportOptionsIdFieldWorksheet(&$worksheet, &$database, $currentLanguageId) {

		// Set the column widths
		// Normal heading
		$heading = array("option_id", "option_name", "option_value_id", "option_value_name");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		$data = array();

		$data[] = $database->query("SELECT option_id, name FROM " . DB_PREFIX . "option_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$data[] = $database->query("SELECT option_value_id, name FROM " . DB_PREFIX . "option_value_description WHERE language_id = '" . $currentLanguageId . "'")->rows;

		$this->exportIdFieldWorksheet($worksheet, $data);
	}

	function exportIdFieldWorksheet(&$worksheet, $data = array()) {

		// Export data info to sheet
		foreach ($data as $key => $val) {
			$startRow = 2;
			foreach ($val as $k => $v) {
				$arrayKeys = array_keys($v);
				$this->setCell($worksheet, $startRow, $key*2, $v[$arrayKeys[0]]);
				$this->setCell($worksheet, $startRow, $key*2+1, htmlspecialchars_decode($v[$arrayKeys[1]]));
				$startRow++;
			}
		}
	}

	function exportOrdersLocationWorksheet(&$worksheet, &$database, $data) {

		$start_date = $data['start_date'] ? $data['start_date'] : '';
		$end_date = $data['end_date'] ? date('Y-m-d', strtotime('+1 day', strtotime($data['end_date']))) : '';

		// Set the column widths
		// Normal heading
		$heading = array("order_id", "order_status", "customer_name", "country_name", "zone_name", "city_name", "county_name", "address");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		// Get current language id
		$currentLanguageId = (int)$this->config->get("config_language_id");

		$data = array();

		$sql = "SELECT o.order_id, os.name as order_status, a.fullname as customer_name, country.name as country_name, zone.name as zone_name, city.name as city_name, county.name as county_name, a.address FROM " . DB_PREFIX . "order o LEFT JOIN " . DB_PREFIX . "order_status os ON (o.order_status_id = os.order_status_id) LEFT JOIN " . DB_PREFIX . "customer c ON (o.customer_id = c.customer_id) LEFT JOIN " . DB_PREFIX . "address a ON (a.address_id = c.address_id) LEFT JOIN " . DB_PREFIX . "country country ON (a.country_id = country.country_id) LEFT JOIN " . DB_PREFIX . "zone zone ON (a.zone_id = zone.zone_id) LEFT JOIN " . DB_PREFIX . "city city ON (a.city_id = city.city_id) LEFT JOIN " . DB_PREFIX . "county county ON (a.county_id = county.county_id) WHERE o.order_status_id != 0 AND os.language_id = " . $currentLanguageId;

		if ($start_date) {
			$sql .= " AND o.date_modified > '" . $start_date . "'";
		}

		if ($end_date) {
			$sql .= " AND o.date_modified < '" . $end_date . "'";
		}

		// echo $sql;exit;

		$data = $database->query($sql)->rows;

		// Export data info to sheet
		$startRow = 2;
		foreach ($data as $key => $val) {
			$startCol = 0;
			foreach ($val as $k => $v) {
				$this->setCell($worksheet, $startRow, $startCol, $v);
				$startCol++;
			}
			$startRow++;
		}
	}

	function exportSalesOrderWorksheet(&$worksheet, &$database, $data) {

		$start_date = $data['start_date'] ? $data['start_date'] : '';
		$end_date = $data['end_date'] ? date('Y-m-d', strtotime('+1 day', strtotime($data['end_date']))) : '';

		// Set the column widths
		// Normal heading
		$heading = array("order_id", "order_status", "products_name", "customer_name", "order_total");
		foreach ($heading as $key => $val) {
			$worksheet->getColumnDimensionByColumn($key)->setWidth(strlen($val) + 1);
			$this->setcell($worksheet, 1, $key, $val);
		}

		// Set sheet title height
		$worksheet->getRowDimension(1)->setRowHeight(30);

		// Get current language id
		$currentLanguageId = (int)$this->config->get("config_language_id");

		$data = array();

		$sql = "SELECT o.order_id, os.name as order_status, pd.name as products_name, o.fullname as customer_name, o.total as order_total FROM " . DB_PREFIX . "order o LEFT JOIN " . DB_PREFIX . "order_status os ON (o.order_status_id = os.order_status_id) LEFT JOIN " . DB_PREFIX . "order_product op ON (o.order_id = op.order_id) LEFT JOIN " . DB_PREFIX . "product_description pd ON (op.product_id = pd.product_id) WHERE o.order_status_id != 0 AND pd.language_id = " . $currentLanguageId . " AND os.language_id = " . $currentLanguageId;

		if ($start_date) {
			$sql .= " AND o.date_modified > '" . $start_date . "'";
		}

		if ($end_date) {
			$sql .= " AND o.date_modified < '" . $end_date . "'";
		}

		$data = $database->query($sql)->rows;

		// print_r($data);exit;

		// Export data info to sheet
		$startRow = 2;
		$orderIdFlag = 0;
		foreach ($data as $key => $val) {
			$startCol = 0;
			if ($val['order_id'] != $orderIdFlag) {
				foreach ($val as $k => $v) {
					$this->setCell($worksheet, $startRow, $startCol, $v);
					$startCol++;
				}
			} else {
				foreach ($val as $k => $v) {
					if ($k == 'products_name') {
						$this->setCell($worksheet, $startRow, $startCol, $v);
					}
					$startCol++;
				}
			}
			$startRow++;
			$orderIdFlag = $val['order_id'];
		}
	}

	function clearCache() {

		$this->cache->delete('*');
	}

	protected function clearSpreadsheetCache() {

		$files = glob(DIR_CACHE . 'Spreadsheet_Excel_Writer' . '*');
		
		if ($files) {
			foreach ($files as $file) {
				if (file_exists($file)) {
					@unlink($file);
					clearstatcache();
				}
			}
		}
	}
}

// Error Handler
function error_handler_for_export($errno, $errstr, $errfile, $errline) {
	
	global $registry;
	
	switch ($errno) {
		case E_NOTICE:
		case E_USER_NOTICE:
			$errors = "Notice";
			break;
		case E_WARNING:
		case E_USER_WARNING:
			$errors = "Warning";
			break;
		case E_ERROR:
		case E_USER_ERROR:
			$errors = "Fatal Error";
			break;
		default:
			$errors = "Unknown";
			break;
	}
	
	$config = $registry->get('config');
	$url = $registry->get('url');
	$request = $registry->get('request');
	$session = $registry->get('session');
	$log = $registry->get('log');
	
	if ($config->get('config_error_log')) {
		$log->write('PHP ' . $errors . ':  ' . $errstr . ' in ' . $errfile . ' on line ' . $errline);
	}

	if (($errors=='Warning') || ($errors=='Unknown')) {
		return true;
	}

	if (($errors != "Fatal Error") && isset($request->get['route']) && ($request->get['route']!='tool/excelproductsimport/download'))  {
		if ($config->get('config_error_display')) {
			echo '<b>' . $errors . '</b>: ' . $errstr . ' in <b>' . $errfile . '</b> on line <b>' . $errline . '</b>';
		}
	} else {
		$session->data['export_error'] = array( 'errstr'=>$errstr, 'errno'=>$errno, 'errfile'=>$errfile, 'errline'=>$errline );
		$token = $request->get['token'];
		$link = $url->link( 'tool/excelproductsimport', 'token='.$token, 'SSL' );
		header('Status: ' . 302);
		header('Location: ' . str_replace(array('&amp;', "\n", "\r"), array('&', '', ''), $link));
		exit();
	}

	return true;
}


function fatal_error_shutdown_handler_for_export() {

	$last_error = error_get_last();
	if ($last_error['type'] === E_ERROR) {
		// fatal error
		error_handler_for_export(E_ERROR, $last_error['message'], $last_error['file'], $last_error['line']);
	}
}