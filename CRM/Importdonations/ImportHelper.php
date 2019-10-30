<?php

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

require_once __DIR__ . '/../../PhpSpreadsheet/vendor/autoload.php';

class CRM_Importdonations_ImportHelper {
  private $logTable = 'viva_salud_import_log';
  private $winbooksFinancialType = 0;
  private $sheetHeader = [];

  public function __construct() {
    $this->checkConfig();
  }

  public function importAnayliticalCodes($excelFile) {
    // open the Excel file, but only the sheet with analytical codes
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setLoadSheetsOnly(['liste analytiques']);
    $spreadsheet = $reader->load($excelFile);

    // store a reference to the sheets
    $worksheetAnalytics = $spreadsheet->getSheetByName('liste analytiques');

    // read column headers header
    $this->readColumnHeader($worksheetAnalytics, 'analytiques');
  }

  public function checkDonors($excelFile) {
    // open the Excel file, but only the sheet with analytical codes
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setLoadSheetsOnly(['donateurs']);
    $spreadsheet = $reader->load($excelFile);

    // store a reference to the sheets
    $worksheetDonors = $spreadsheet->getSheetByName('donateurs');

    // read column headers header
    $this->readColumnHeader($worksheetDonors, 'donateurs');

    // validate the sheets
    $this->validateSheetHeader($worksheetDonors, 'donateurs', ['NUMBER', 'ADRESS1']);

    // check donateurs
    $itemsProcessed = $this->logDonorDifferences($worksheetDonors);

    return $itemsProcessed;
  }

  public function importDonations($excelFile) {
    try {
      // open the Excel file, and open only the sheets we're interested in: donteurs + transit
      $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
      $reader->setLoadSheetsOnly(['donateurs', 'transit', 'liste analytiques']);
      $spreadsheet = $reader->load($excelFile);

      // store a reference to the sheets
      $worksheetDonateurs = $spreadsheet->getSheetByName('donateurs');
      $worksheetTransit = $spreadsheet->getSheetByName('transit');

      // read column headers header
      $this->readColumnHeader($worksheetDonateurs, 'donateurs');
      $this->readColumnHeader($worksheetTransit, 'transit');

      // validate the sheets
      $this->validateSheetHeader($worksheetDonateurs,  ['NUMBER', 'ADRESS1']);
      $this->validateSheetHeader($worksheetTransit, ['DATE', 'COMMENT', 'NAME', 'AMOUNTEUR', 'ACCOUNTGL']);

      // delete donations within the range of the worksheet
      $this->deleteExistingDonations($worksheetTransit);

      // import transit
      $this->importTransit($worksheetTransit);

      CRM_Core_Session::setStatus('OK', 'Success', 'success');
    }
    catch (Exception $e) {
      //CRM_Core_Error::fatal($e->getMessage());
      CRM_Core_Session::setStatus($e->getMessage(), '', 'error');
    }
  }

  /**
   * @param Worksheet $worksheet
   */
  private function importTransit($worksheet) {
    $i = 2;
    while (($winbooksCode = trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['Trs(ZONANA5)'], $i))->getFormattedValue()) != '') {
      // make sure we have a value in the column "comment"
      if (trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['COMMENT'], $i)->getFormattedValue()) != '') {
        // lookup the contact
        $params = [
          'external_identifier' => $winbooksCode,
          'sequential' => 1,
        ];
        $contact = civicrm_api3('Contact', 'get', $params);
        if ($contact['count'] > 0) {
          $date = trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['DATE'], $i)->getFormattedValue());
          // convert to YYYY-MM-DD
          $dateParts = explode('/', $date);
          $formattedDate = $dateParts[2] . '-' . sprintf("%02d", $dateParts[0]) . '-' . sprintf("%02d", $dateParts[1]);

          $params = [
            'contact_id' => $contact['values'][0]['id'],
            'source' => trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['NAME'], $i)->getFormattedValue()),
            'total_amount' => str_replace('-', '', trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['AMOUNTEUR'], $i)->getFormattedValue())),
            'receive_date' => $formattedDate,
            'contribution_status_id' => 1, // completed
            'financial_type_id' => $this->winbooksFinancialType,
          ];
          civicrm_api3('Contribution', 'create', $params);
        }
        else {
          $this->logComment('transit', "line $i", 'Donation not imported', "Contact $winbooksCode does not exist in CiviCRM");
        }
      }

      $i++;
    }
  }

  /**
   * @param Worksheet $worksheet
   */
  private function logDonorDifferences($worksheet) {
    $i = 3; // TODO: IS THIS ALWAYS LIKE THIS?????

    while (($winbooksCode = $this->getCellValueByColName($worksheet, 'donateurs', $i, 'NUMBER')) != '') {
      // lookup the contact
      $params = [
        'external_identifier' => $winbooksCode,
        'sequential' => 1,
      ];
      $contact = civicrm_api3('Contact', 'get', $params);
      if ($contact['count'] > 0) {
        // check the address
        $street = $this->getCellValueByColName($worksheet, 'donateurs', $i, 'ADRESS1');
        $sql = "select count(*) from civicrm_address where replace(replace(street_address, ',', ''), ' ', '') = replace(replace(%1, ',', ''), ' ', '') and contact_id = %2";
        $sqlParams = [
          1 => [$street, 'String'],
          2 => [$contact['values'][0]['id'], 'Integer'],
        ];
        $n = CRM_Core_DAO::singleValueQuery($sql, $sqlParams);
        if ($n == 0) {
          $this->logComment('donateurs', "line $i", 'Address not found in CiviCRM', "$winbooksCode, " . $this->getCellValueByColName($worksheet, 'donateurs', $i, 'ADRESS1'));
        }
      }
      else {
        $this->logComment('donateurs', "line $i", 'Contact not found in CiviCRM', "$winbooksCode, " . $this->getCellValueByColName($worksheet, 'donateurs', $i, 'NUMBER'));
      }

      $i++;
    }

    // return the number of processed items
    return $i - 2;
  }

  /**
   * @param Worksheet $worksheet
   */
  private function deleteExistingDonations($worksheet) {
    // find the lowest and highest date
    $lowestDate = '3000-01-01';
    $highestDate = '1000-01-01';
    $i = 2;
    while (($date = trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['ACCOUNTGL'], $i)->getFormattedValue())) != '') {
      // convert to YYYY-MM-DD
      $dateParts = explode('/', $date);
      $formattedDate = $dateParts[2] . '-' . sprintf("%02d", $dateParts[0]) . '-' . sprintf("%02d", $dateParts[1]);

      // make sure we have a value in the column "comment"
      if (trim($worksheet->getCellByColumnAndRow($this->sheetHeader['transit']['DATE'], $i)->getFormattedValue()) != '') {
        if ($formattedDate < $lowestDate) {
          $lowestDate = $formattedDate;
        }
        if ($formattedDate > $highestDate) {
          $highestDate = $formattedDate;
        }
      }
      $i++;
    }

    // delete all contributions of type "winbooks" between these dates
    $sql = "
      delete from 
        civicrm_contribution 
      where
        receive_date between '$lowestDate 00:00' and '$highestDate 23:59'
      and 
        financial_type_id = {$this->winbooksFinancialType}
    ";
    CRM_Core_DAO::executeQuery($sql);
  }

  /**
   * Store the column headers names and position, so we don't depend on the exact order/number of columns
   *
   * @param Worksheet $worksheet
   * @param String $sheetName
   */
  private function readColumnHeader($worksheet, $sheetName) {
    $this->sheetHeader[$sheetName] = [];

    $i = 1;
    while (($c = $this->getCellValue($worksheet, 1, $i)) != '') {
      $this->sheetHeader[$sheetName][$c] = $i;
      $i++;
    }
  }

  /**
   * @param Worksheet $worksheet
   */
  private function validateSheetHeader($worksheet, $worksheetName, $expectedColumns) {
    foreach ($expectedColumns as $expectedColumn) {
      if (!array_key_exists($expectedColumn, $this->sheetHeader[$worksheetName])) {
        throw new Exception("Expected a column $expectedColumn in worksheet $worksheetName");
      }
    }
  }

  private function logComment($worksheetName, $cell, $commentType, $comment) {
    $sql = "insert into {$this->logTable} (worksheet, cell, comment_type, comment) values (%1, %2, %3, %4)";
    $sqlParams = [
      1 => [$worksheetName, 'String'],
      2 => [$cell, 'String'],
      3 => [$commentType, 'String'],
      4 => [$comment, 'String'],
    ];
    CRM_Core_DAO::executeQuery($sql, $sqlParams);
  }

  private function checkConfig() {
    // make sure the log table exists
    $sql = "
      CREATE TABLE IF NOT EXISTS {$this->logTable} (
          id int(10) unsigned NOT NULL AUTO_INCREMENT,
          worksheet varchar(255),
          cell varchar(32),
          comment_type varchar(255),
          comment varchar(255),
          PRIMARY KEY (`id`)
      )
    ";
    CRM_Core_DAO::executeQuery($sql);

    // clear all records
    $sql = "truncate table {$this->logTable}";
    CRM_Core_DAO::executeQuery($sql);

    // make sure the contrubution type "winbooks" exists
    $params = [
      'sequential' => 1,
      'name' => 'Donation (Winbooks)',
    ];
    try {
      $ft = civicrm_api3('FinancialType', 'getsingle', $params);
      $this->winbooksFinancialType = $ft['id'];
    }
    catch (Exception $e) {
      // doesn't exist, create it
      $params['is_active'] = 1;
      $params['is_deductible'] = 0;
      $ft = civicrm_api3('FinancialType', 'create', $params);
      $this->winbooksFinancialType = $ft['id'];
    }
  }

  private function getCellValue($worksheet, $row, $col) {
    // return trimmed cell value
    return trim($worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue());
  }

  private function getCellValueByColName($worksheet, $worksheetName, $row, $colName) {
    return $this->getCellValue($worksheet, $row, $this->sheetHeader[$worksheetName][$colName]);
  }
}
