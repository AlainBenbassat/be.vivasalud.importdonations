<?php

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

require_once __DIR__ . '/../../PhpSpreadsheet/vendor/autoload.php';

class CRM_Importdonations_ImportHelper {
  private $logTable = 'viva_salud_import_log';
  private $winbooksFinancialType = 0;

  public function __construct() {
    $this->checkConfig();
  }

  public function import($excelFile) {
    try {
      // open the Excel file, and open only the sheets we're interested in: donteurs + transit
      $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
      $reader->setLoadSheetsOnly(['donateurs', 'transit']);
      $spreadsheet = $reader->load($excelFile);

      // get the "donateurs" and "transit" worksheets
      $worksheetDonateurs = $spreadsheet->getSheetByName('donateurs');
      $worksheetTransit = $spreadsheet->getSheetByName('transit');

      // validate the sheets
      $this->validateSheetHeader($worksheetDonateurs,  ['', 'NUMBER', 'NAME1', 'NAME2', 'ADRESS1', 'ADRESS2', 'VATCAT', 'COUNTRY', 'VATNUMBER', 'ZIPCODE', 'CITY', 'LANG', 'CATEGORY', 'TOTDEB2', 'TOTCRE2', 'SOLDE', 'REPORT', 'IBANAUTO', 'BICAUTO']);
      $this->validateSheetHeader($worksheetTransit, ['', 'ACCOUNT', 'NAME', 'DBKCODE', 'DBKTYPE', 'DOCNUMBER', 'BOOKYEAR', 'PERIOD', 'ACCOUNTGL', 'DATE', 'COMMENT', 'AMOUNTEUR', 'MATCHNO', 'OLDDATE', 'ISMATCHED', 'ISLOCKED', 'ISIMPORTED', 'ISPOSITIVE', 'ISTEMP', 'MEMOTYPE', 'ISDOC', 'LINEORDER', 'AMOUNTGL', 'Fin(ZONANA1)', 'Act(ZONANA2)', 'Mvt(ZONANA3)', 'Mdp(ZONANA4)', 'Trs(ZONANA5)', 'Att(ZONANA6)']);

      // check donateurs
      $this->checkDonateurs($worksheetDonateurs);

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
    while (($date = $worksheet->getCell("AA$i")) != '') {
      // make sure we have a value in the column "comment"
      if (trim($worksheet->getCell("J$i")) != '') {
        // get the winbooks client code
        $winbooksCode = trim($worksheet->getCell("AA$i"));

        // lookup the contact
        $params = [
          'external_identifier' => $winbooksCode,
          'sequential' => 1,
        ];
        $contact = civicrm_api3('Contact', 'get', $params);
        if ($contact['count'] > 0) {
          $date = $worksheet->getCell("I$i")->getFormattedValue();
          // convert to YYYY-MM-DD
          $dateParts = explode('/', $date);
          $formattedDate = $dateParts[2] . '-' . sprintf("%02d", $dateParts[0]) . '-' . sprintf("%02d", $dateParts[1]);

          $params = [
            'contact_id' => $contact['values'][0]['id'],
            'source' => trim($worksheet->getCell("B$i")),
            'total_amount' => str_replace('-', '', $worksheet->getCell("K$i")),
            'receive_date' => $formattedDate,
            'contribution_status_id' => 1, // completed
            'financial_type_id' => $this->winbooksFinancialType,
          ];
          civicrm_api3('Contribution', 'create', $params);
        }
        else {
          $this->logComment('transit', "AA$i", 'Donation not imported', "Contact $winbooksCode does not exist in CiviCRM");
        }
      }

      $i++;
    }

  }

  /**
   * @param Worksheet $worksheet
   */
  private function checkDonateurs($worksheet) {
    $i = 2;
    while (($winbooksCode = trim($worksheet->getCell("A$i"))) != '') {
      // lookup the contact
      $params = [
        'external_identifier' => $winbooksCode,
        'sequential' => 1,
      ];
      $contact = civicrm_api3('Contact', 'get', $params);
      if ($contact['count'] > 0) {
        // check the address
        $street = trim($worksheet->getCell("D$i"));
        $sql = "select count(*) from civicrm_address where replace(replace(street_address, ',', ''), ' ', '') = replace(replace(%1, ',', ''), ' ', '') and contact_id = %2";
        $sqlParams = [
          1 => [$street, 'String'],
          2 => [$contact['values'][0]['id'], 'Integer'],
        ];
        $n = CRM_Core_DAO::singleValueQuery($sql, $sqlParams);
        if ($n == 0) {
          $this->logComment('donateurs', "A$i", 'Address not found in CiviCRM', "$winbooksCode, " . trim($worksheet->getCell("B$i")) . ', ' . trim($worksheet->getCell("D$i")));
        }
      }
      else {
        $this->logComment('donateurs', "A$i", 'Contact not found in CiviCRM', "$winbooksCode, " . trim($worksheet->getCell("B$i")));
      }

      $i++;
    }
  }

  /**
   * @param Worksheet $worksheet
   */
  private function deleteExistingDonations($worksheet) {
    // find the lowest and highest date
    $lowestDate = '3000-01-01';
    $highestDate = '1000-01-01';
    $i = 2;
    while (($date = $worksheet->getCell("I$i")->getFormattedValue()) != '') {
      // convert to YYYY-MM-DD
      $dateParts = explode('/', $date);
      $formattedDate = $dateParts[2] . '-' . sprintf("%02d", $dateParts[0]) . '-' . sprintf("%02d", $dateParts[1]);

      // make sure we have a value in the column "comment"
      if (trim($worksheet->getCell("J$i")) != '') {
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
   * @param Worksheet $worksheet
   */
  private function validateSheetHeader($worksheet, $expectedColumns) {
    for ($i = 1; $i < count($expectedColumns); $i++) {
      if ($worksheet->getCellByColumnAndRow($i, 1) != $expectedColumns[$i]) {
        throw new Exception("Expected column $i to be " . $expectedColumns[$i] . ' but found ' . $worksheet->getCellByColumnAndRow($i, 1));
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
}
