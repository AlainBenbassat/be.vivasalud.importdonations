<?php
use CRM_Importdonations_ExtensionUtil as E;

class CRM_Importdonations_Page_ImportLog extends CRM_Core_Page {

  public function run() {
    CRM_Utils_System::setTitle('Import Donations - Status');

    // get the log
    $dao = CRM_Core_DAO::executeQuery("select * from viva_salud_import_log order by worksheet, comment_type");
    $log = $dao->fetchAll();
    $this->assign('log', $log);

    parent::run();
  }

}
