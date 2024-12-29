<?php
// import all MS Word default styles and use them to add a new content

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->importStylesWordDefault();

$docx->addText('This is the resulting paragraph with the default "Heading1" style.', array('pStyle' => 'Heading1'));

$valuesTable = array(
    array(11, 12, 13, 14),
    array(21, 22, 23, 24),
    array(31, 32, 33, 34),
);
$paramsTable = array(
    'tableStyle' => 'TableGrid',
);
$docx->addTable($valuesTable, $paramsTable);

$docx->createDocx('example_importStylesWordDefault_1');