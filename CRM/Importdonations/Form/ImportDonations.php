<?php

use CRM_Importdonations_ExtensionUtil as E;

class CRM_Importdonations_Form_ImportDonations extends CRM_Core_Form {
  public function buildQuickForm() {
    CRM_Utils_System::setTitle('Viva Salud - import donations from accounting Excel');

    $this->add('File', 'uploadFile', 'Account Excel file<br>(yyyymmdd yyyy rapports des dons.xlsx)', 'size=30 maxlength=255', TRUE);

    $this->addButtons([
      [
        'type' => 'submit',
        'name' => 'Import',
        'isDefault' => TRUE,
      ],
      [
        'type' => 'cancel',
        'name' => 'Cancel',
      ],
    ]);

    // export form elements
    $this->assign('elementNames', $this->getRenderableElementNames());
    parent::buildQuickForm();
  }

  public function cancelAction() {
    // redirect to the main page
    CRM_Utils_System::redirect(CRM_Utils_System::url('civicrm', ''));
  }

  public function postProcess() {
    $values = $this->exportValues();

    // get the selected file
// TIJDELIJK!!!!!!!!!!!!!!!!    $tmpFileName = $this->_submitFiles['uploadFile']['tmp_name'];
    $tmpFileName = '/home/alain/Documents/Klanten/Viva Salud/winbooks import/20190926 2019 rapports des dons.xlsx';

    if (!$tmpFileName) {
      CRM_Core_Session::setStatus('Cannot open ' . $this->_submitFiles['uploadFile']['name'] . '. Maybe it\'s too big?', 'Error', 'error');
    }
    else {
      // import the transactions
      try {
        $importHelper = new CRM_Importdonations_ImportHelper();
        $importHelper->import($tmpFileName);
      }
      catch (Exception $e) {
        CRM_Core_Session::setStatus($e->getMessage(), 'Import', 'error');
      }
    }

    parent::postProcess();
  }


  public function getRenderableElementNames() {
    $elementNames = array();
    foreach ($this->_elements as $element) {
      $label = $element->getLabel();
      if (!empty($label)) {
        $elementNames[] = $element->getName();
      }
    }
    return $elementNames;
  }

}
