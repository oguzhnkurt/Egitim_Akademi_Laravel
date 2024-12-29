<?php
// modify page layout to A4-landscape and 2 columns

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocxFromTemplate('../../files/MultipleSections.docx');

// using the sectionNumbers option one may choose the sections to modify
$docx->modifyPageLayout('A4-landscape', array('numberCols' => '2', 'sectionNumbers' => array(2)));

$docx->createDocx('example_modifyPageLayout_2');