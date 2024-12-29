<?php
// import all MS Word default styles into a DOCX template and use them to add a new content

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocxFromTemplate('../../files/TemplateSimpleText.docx');

// 'ignore' (default) and 'replace' import types can be used to handle existing styles in the DOCX template
$docx->importStylesWordDefault();
// $docx->importStylesWordDefault('replace');

$docx->addText('This is the resulting paragraph with the default "Heading1" style.', array('pStyle' => 'Heading1'));

$docx->createDocx('example_importStylesWordDefault_3');