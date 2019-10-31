<?php

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

require_once __DIR__ . '/../../PhpSpreadsheet/vendor/autoload.php';

class CRM_Importdonations_ImportHelper {
  private $logTable = 'viva_salud_import_log';
  private $winbooksFinancialType = 0;
  private $optionGroupMdp = 0;
  private $optionGroupFin = 162;
  private $optionGroupAct = 163;
  private $optionGroupMvt = 164;
  private $optionGroupAtt = 165;
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

    // validate the sheets
    $this->validateSheetHeader($worksheetAnalytics, 'analytiques', ['Section', 'Referency', 'Name']);

    // sync Act, Mvt, Fin, Att lists with their corresponding CiviCRM option group
    $actNum = $this->addOptionGroupValues($worksheetAnalytics, 'analytiques', 'Act', $this->optionGroupAct);
    $mvtNum = $this->addOptionGroupValues($worksheetAnalytics, 'analytiques', 'Mvt', $this->optionGroupMvt);
    $finNum = $this->addOptionGroupValues($worksheetAnalytics, 'analytiques', 'Fin', $this->optionGroupFin);
    $attNum = $this->addOptionGroupValues($worksheetAnalytics, 'analytiques', 'Att', $this->optionGroupAtt);
    $mdpNum = $this->addOptionGroupValues($worksheetAnalytics, 'analytiques', 'Mdp', $this->optionGroupMdp);

    return $actNum + $mvtNum + $finNum + $attNum + $mdpNum;
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
    // open the Excel file, and open only the sheets we're interested in: donteurs + transit
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setLoadSheetsOnly(['donateurs', 'transit']);
    $spreadsheet = $reader->load($excelFile);

    // store a reference to the sheets
    $worksheetDonors = $spreadsheet->getSheetByName('donateurs');
    $worksheetTransit = $spreadsheet->getSheetByName('transit');

    // read column headers header
    $this->readColumnHeader($worksheetDonors, 'donateurs');
    $this->readColumnHeader($worksheetTransit, 'transit');

    // validate the sheets
    $this->validateSheetHeader($worksheetDonors,  ['NUMBER', 'ADRESS1']);
    $this->validateSheetHeader($worksheetTransit, ['DATE', 'COMMENT', 'NAME', 'AMOUNTEUR', 'ACCOUNTGL']);

    // delete donations within the range of the worksheet
    $this->deleteExistingDonations($worksheetTransit);

    // import transit
    $itemsProcessed = $this->importTransit($worksheetTransit, $worksheetDonors);

    return $itemsProcessed;
  }

  /**
   * @param Worksheet $worksheetTransit
   * @param Worksheet $worksheetDonors
   */
  private function importTransit($worksheetTransit, $worksheetDonors) {
    $i = 2;
    while (($winbooksCode = $this->getCellValueByColName($worksheetTransit, 'transit', $i, 'Trs(ZONANA5)')) != '') {
      // make sure we have a value in the column "comment"
      if ($this->getCellValueByColName($worksheetTransit, 'transit', $i, 'COMMENT') != '') {
        // lookup the contact
        $params = [
          'external_identifier' => $winbooksCode,
          'sequential' => 1,
        ];
        $contact = civicrm_api3('Contact', 'get', $params);
        if ($contact['count'] == 0) {
          // create the contact
          $contactID = $this->createContact($worksheetDonors, $winbooksCode);
        }
        else {
          $contactID = $contact['values'][0]['id'];
        }

          $date = $this->getCellValueByColName($worksheetTransit, 'transit', $i, 'DATE');
          // convert to YYYY-MM-DD
          $dateParts = explode('/', $date);
          $formattedDate = $dateParts[2] . '-' . sprintf("%02d", $dateParts[0]) . '-' . sprintf("%02d", $dateParts[1]);

          $params = [
            'contact_id' => $contactID,
            'source' => $this->getCellValueByColName($worksheetTransit, 'transit', $i, 'NAME'),
            'total_amount' => str_replace('-', '', $this->getCellValueByColName($worksheetTransit, 'transit', $i, 'AMOUNTEUR')),
            'receive_date' => $formattedDate,
            'contribution_status_id' => 1, // completed
            'financial_type_id' => $this->winbooksFinancialType,
          ];
          civicrm_api3('Contribution', 'create', $params);
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
          $this->logComment('donateurs', "line $i", 'Address not found in CiviCRM', "$winbooksCode, " . $street);
        }
      }
      else {
        $this->logComment('donateurs', "line $i", 'Contact not found in CiviCRM', "$winbooksCode, " . $this->getCellValueByColName($worksheet, 'donateurs', $i, 'NAME1'). ", " . $this->getCellValueByColName($worksheet, 'donateurs', $i, 'ADRESS1'));
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
    while (($date = $this->getCellValueByColName($worksheet, 'transit', $i, 'ACCOUNTGL')) != '') {
      // convert to YYYY-MM-DD
      $dateParts = explode('/', $date);
      $formattedDate = $dateParts[2] . '-' . sprintf("%02d", $dateParts[0]) . '-' . sprintf("%02d", $dateParts[1]);

      // make sure we have a value in the column "comment"
      if ($this->getCellValueByColName($worksheet, 'transit', $i, 'DATE') != '') {
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

  /**
   * @param Worksheet $worksheetAnalytics
   * @param $worksheetName
   * @param $section
   * @param $optionGroupID
   */
  private function addOptionGroupValues($worksheet, $worksheetName, $section, $optionGroupID) {
    $i = 2;
    $created = 0;
    while (($excelSection = $this->getCellValueByColName($worksheet, $worksheetName, $i, 'Section')) != '') {
      if ($excelSection == $section) {
        $code = $this->getCellValueByColName($worksheet, $worksheetName, $i, 'Referency');
        $name = $this->getCellValueByColName($worksheet, $worksheetName, $i, 'Name');

        // find the corresponding option value
        $params = [
          'sequential' => 1,
          'option_group_id' => $optionGroupID,
          'value' => $code,
        ];
        $result = civicrm_api3('OptionValue', 'get', $params);
        if ($result['count'] == 0) {
          // add the value
          $params['label'] = $name;
          $result = civicrm_api3('OptionValue', 'create', $params);
          $created++;

          // log this addition
          $this->logComment($worksheetName, "line $i", 'Option Value added in CiviCRM', "$section, $code, $name");
        }
      }

      $i++;
    }

    return $created;
  }

  private function createContact($worksheet, $winbooksCode) {
    // lookup the contact in the donor list
    $found = FALSE;
    $i = 3;
    while (($excelWinbooksCode = $this->getCellValueByColName($worksheet, 'donateurs', $i, 'NUMBER')) != '') {
      if ($excelWinbooksCode == $winbooksCode) {
        $found = TRUE;
        break;
      }
    }

    if ($found == FALSE) {
      throw new Exception("Cannot find $winbooksCode in 'donateurs'");
    }

    $params = [
      'sequential' => 1,
    ];

    // determine the contact type
    $contactPrefix = $this->getCellValueByColName($worksheet, 'donateurs', $i, 'CIVNAME1');
    $contactType = $this->getContactType($contactPrefix);
    $params['contact_type'] = $contactType;

    // get the name
    $contactName = $this->getCellValueByColName($worksheet, 'donateurs', $i, 'NUMBER');
    if ($contactType == 'Individual') {
      $names = $this->getFirstNameLastName($contactName);
      $params['first_name'] = $names['first_name'];
      $params['last_name'] = $names['last_name'];
    }
    elseif ($contactType == 'Organization') {
      $params['organization_name'] = $contactName;
    }
    else {
      $params['household_name'] = $contactName;
    }

    // add preferred language
    $lang = $this->getCellValueByColName($worksheet, 'donateurs', $i, 'LANG');
    if ($lang == 'N') {
      $params['preferred_language'] = 'nl_NL';
    }
    elseif ($lang == 'F') {
      $params['preferred_language'] = 'fr_FR';
    }
    elseif ($lang == 'E') {
      $params['preferred_language'] = 'en_US';
    }

    // add address


    return $contactID;
  }

  private function getFirstNameLastName($contactName) {
    $names = [];

    // try to extract first and last name: we assume the last part is the first name
    $nameParts = explode(' ', $contactName);
    if (count($nameParts) > 1) {
      $names['first_name'] = $nameParts[count($nameParts) - 1];
      unset($nameParts[count($nameParts) - 1]);
      $names['last_name'] = implode(' ', $nameParts);
    }
    else {
      // just one "word", assume it's the last name
      $names['first_name'] = '-';
      $names['last_name'] = $contactName;
    }

    return $names;
  }

  private function getContactType($contactPrefix) {
    switch ($contactPrefix) {
      case 'MEVR':
      case 'MME':
      case 'MRS':
      case 'DHR':
      case 'M':
      case 'MR':
        $retval = 'Individual';
        break;
      case 'ASBL':
      case 'VZW':
      case 'SPRL':
      case 'BVBA':
      case 'OND':
        $retval = 'Organization';
        break;
      default:
        $retval = 'Household';
    }

    return $retval;
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

    // make sure the option group "payment method" exists
    $params = [
      'sequential' => 1,
      'name' => 'payment_method_accounting',
    ];
    try {
      $og = civicrm_api3('OptionGroup', 'getsingle', $params);
      $this->optionGroupMdp = $og['id'];
    }
    catch (Exception $e) {
      // doesn't exist, create it
      $params['title'] = 'Betaalmethode';
      $og = civicrm_api3('OptionGroup', 'create', $params);
      $this->optionGroupMdp = $og['id'];
    }

  }

  private function getCellValue($worksheet, $row, $col) {
    // return trimmed cell value
    return trim($worksheet->getCellByColumnAndRow($col, $row)->getFormattedValue());
  }

  private function getCellValueByColName($worksheet, $worksheetName, $row, $colName) {
    // make sure the column name exists
    if (isset($this->sheetHeader[$worksheetName][$colName])) {
      return $this->getCellValue($worksheet, $row, $this->sheetHeader[$worksheetName][$colName]);
    }
    else {
      throw new Exception("Column $colName not found in $worksheetName");
    }
  }
}
