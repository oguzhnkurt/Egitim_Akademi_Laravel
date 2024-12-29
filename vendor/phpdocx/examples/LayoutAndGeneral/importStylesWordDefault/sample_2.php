<?php
// import specific MS Word default styles and use them to add a new content

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->importStylesWordDefault('ignore', array('Heading1', 'Heading2', 'Heading3'));

$docx->addText('Heading 1.', array('pStyle' => 'Heading1'));
$docx->addText('Heading 2.', array('pStyle' => 'Heading2'));
$docx->addText('Heading 3.', array('pStyle' => 'Heading3'));

$docx->createDocx('example_importStylesWordDefault_2');