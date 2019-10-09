<?php

require_once __DIR__ . '/../../PhpSpreadsheet/vendor/autoload.php';

class CRM_Importdonations_ImportHelper {
  public function import($excelFile) {
    try {
      $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($excelFile);
      CRM_Core_Session::setStatus('OK', 'Success', 'success');
    }
    catch (Exception $e) {
      CRM_Core_Session::setStatus($e->getMessage(), '', 'error');
    }


  }
}
