<?php 
class ControllerToolExcelImportExport extends Controller { 
	private $error = array();
	
	public function index() {

		$this->load->language('tool/excelimportexport');
		$this->document->setTitle($this->language->get('heading_title'));

		if(!empty($this->session->data['export_error']['errstr'])) {
			$this->error['warning'] = $this->session->data['export_error']['errstr'];
			
			if(!empty($this->session->data['export_nochange'])) {
				$this->error['warning'] .= '<br />'.$this->language->get( 'text_nochange' );
			}
			
			$this->error['warning'] .= '<br />'.$this->language->get( 'text_log_details' );
		}
		unset($this->session->data['export_error']);
		unset($this->session->data['export_nochange']);

		$data['product_import_heading_title'] = $this->language->get('product_import_heading_title');
		$data['sales_export_heading_title'] = $this->language->get('sales_export_heading_title');
		$data['orders_location_heading_title'] = $this->language->get('orders_location_heading_title');

		$data['text_choose_import'] = $this->language->get('text_choose_import');
		$data['text_export_sample'] = $this->language->get('text_export_sample');
		$data['text_export_id_field'] = $this->language->get('text_export_id_field');
		$data['text_export_order_location'] = $this->language->get('text_export_order_location');
		$data['text_start_date'] = $this->language->get('text_start_date');
		$data['text_end_date'] = $this->language->get('text_end_date');

		$data['heading_title'] = $this->language->get('heading_title');
		$data['button_import'] = $this->language->get('button_import');
		$data['button_export'] = $this->language->get('button_export');

		$data['button_export_pid'] = $this->language->get('button_export_pid');
		$data['button_export_page'] = $this->language->get('button_export_page');

		$data['tab_general'] = $this->language->get('tab_general');
		$data['error_select_file'] = $this->language->get('error_select_file');
		$data['error_post_max_size'] = $this->language->get('error_post_max_size');
		$data['error_upload_max_filesize'] = $this->language->get('error_upload_max_filesize');

 		if (isset($this->error['warning'])) {
			$data['error_warning'] = $this->error['warning'];
		} else {
			$data['error_warning'] = '';
		}
		
		if (isset($this->session->data['success'])) {
			$data['success'] = $this->session->data['success'];
		
			unset($this->session->data['success']);
		} else {
			$data['success'] = '';
		}
		
		$data['breadcrumbs'] = array();

		$data['breadcrumbs'][] = array(
			'text'      => $this->language->get('text_home'),
			'href'      => $this->url->link('common/home', 'token=' . $this->session->data['token'], 'SSL'),
			'separator' => FALSE
		);

		$data['breadcrumbs'][] = array(
			'text'      => $this->language->get('heading_title'),
			'href'      => $this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'),
			'separator' => ' :: '
		);
		
		$data['import_products'] = $this->url->link('tool/excelimportexport/importProducts', 'token=' . $this->session->data['token'], 'SSL');
		$data['export_sample'] = $this->url->link('tool/excelimportexport/exportSample', 'token=' . $this->session->data['token'], 'SSL');
		$data['export_id_field'] = $this->url->link('tool/excelimportexport/exportIdField', 'token=' . $this->session->data['token'], 'SSL');
		$data['export_orders_location'] = $this->url->link('tool/excelimportexport/exportOrdersLocation', 'token=' . $this->session->data['token'], 'SSL');
		$data['export_sales_order'] = $this->url->link('tool/excelimportexport/exportSalesOrder', 'token=' . $this->session->data['token'], 'SSL');

		$data['post_max_size'] = $this->return_bytes(ini_get('post_max_size'));
		$data['upload_max_filesize'] = $this->return_bytes(ini_get('upload_max_filesize'));

		$data['header'] = $this->load->controller('common/header');
		$data['column_left'] = $this->load->controller('common/column_left');
		$data['footer'] = $this->load->controller('common/footer');
		
		$this->response->setOutput($this->load->view('tool/excelimportexport.tpl', $data));
	}

	function return_bytes($val)	{

		$val = trim($val);
	
		switch (strtolower(substr($val, -1)))
		{
			case 'm': $val = (int)substr($val, 0, -1) * 1048576; break;
			case 'k': $val = (int)substr($val, 0, -1) * 1024; break;
			case 'g': $val = (int)substr($val, 0, -1) * 1073741824; break;
			case 'b':
				switch (strtolower(substr($val, -2, 1)))
				{
					case 'm': $val = (int)substr($val, 0, -2) * 1048576; break;
					case 'k': $val = (int)substr($val, 0, -2) * 1024; break;
					case 'g': $val = (int)substr($val, 0, -2) * 1073741824; break;
					default : break;
				} break;
			default: break;
		}
		return $val;
	}

	public function importProducts() {

		if(($this->request->server['REQUEST_METHOD'] == 'POST') && ($this->validate())) {
			
			if ((isset($this->request->files['upload'])) && (is_uploaded_file($this->request->files['upload']['tmp_name']))) {
			
				$file = $this->request->files['upload']['tmp_name'];

				$this->load->language('tool/excelimportexport');
				$this->load->model('tool/excelimportexport');

				if ($this->model_tool_excelimportexport->import($file)===TRUE) {
			
					$this->session->data['success'] = $this->language->get('text_product_import_success');
					$this->response->redirect($this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'));
				}else{
					$this->error['warning'] = $this->language->get('error_upload');
					$this->response->redirect($this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'));
				}
			}
		}
	}

	public function exportSample() {

		if ($this->validate()) {
			$this->load->model('tool/excelimportexport');
			$this->model_tool_excelimportexport->exportSample();

			$this->response->redirect($this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'));

		} else {
			// return a permission error page
			return $this->forward('error/permission');
		}
	}

	public function exportIdField() {

		if ($this->validate()) {
			$this->load->model('tool/excelimportexport');
			$this->model_tool_excelimportexport->exportIdField();
			
			$this->response->redirect($this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'));

		} else {
			// return a permission error page
			return $this->forward('error/permission');
		}
	}

	public function exportOrdersLocation() {

		if (($this->request->server['REQUEST_METHOD'] == 'POST') && ($this->validate())) {
			
			$this->load->model('tool/excelimportexport');
			$this->model_tool_excelimportexport->exportOrdersLocation($this->request->post);
			
			$this->response->redirect($this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'));

		} else {
			// return a permission error page
			return $this->forward('error/permission');
		}
	}

	public function exportSalesOrder() {

		if (($this->request->server['REQUEST_METHOD'] == 'POST') && ($this->validate())) {
			$this->load->model('tool/excelimportexport');
			$this->model_tool_excelimportexport->exportSalesOrder($this->request->post);
			
			$this->response->redirect($this->url->link('tool/excelimportexport', 'token=' . $this->session->data['token'], 'SSL'));

		} else {
			// return a permission error page
			return $this->forward('error/permission');
		}
	}

	private function validate() {
		if (!$this->user->hasPermission('modify', 'tool/excelimportexport')) {
			$this->error['warning'] = $this->language->get('error_permission');
		}
		
		if (!$this->error) {
			return TRUE;
		} else {
			return FALSE;
		}
	}
}