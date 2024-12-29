<?php
// remove image variables (placeholders) from an existing DOCX

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocxFromTemplate('../../files/placeholderImage_multiple.docx');

// remove all images with IMAGE_1 as placeholder
$docx->removeTemplateVariableImage('IMAGE_1');

// remove only the first IMAGE_1 placeholder
//$docx->removeTemplateVariableImage('IMAGE_1', 'inline', 'document', array('firstMatch' => true));

$docx->createDocx('example_removeTemplateVariable_1');