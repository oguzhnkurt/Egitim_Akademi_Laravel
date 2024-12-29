<?php
// import styles from an existing DOCX and use them to add a new content

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();
// you may first check the available styles using the parseStyles('../files/TemplateStyles.docx') method

$docx->importStyles('../../files/TemplateStyles.docx', 'merge', array('crazyStyle'));

$docx->addText('This is the resulting paragraph with the "CrazyStyle".', array('pStyle' => 'crazyStyle'));

// you may also import a complete XML style sheet by
// $docx->importStyles('../files/TemplateStyles.docx', 'replace');

$docx->createDocx('example_importStyles_1');