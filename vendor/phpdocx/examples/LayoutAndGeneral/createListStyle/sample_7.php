<?php
// create and apply the same custom list style in different lists

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

// custom options
$latinListOptions = array();
$latinListOptions[0]['type'] = 'lowerLetter';
$latinListOptions[0]['format'] = '%1.';
$latinListOptions[1]['type'] = 'lowerRoman';
$latinListOptions[1]['format'] = '%1.%2.';

// create the list style with name: latin
$docx->createListStyle('latin', $latinListOptions);

// list items
$myList = array('item 1', array('subitem 1.1', 'subitem 1.2'), 'item 2');

// insert the custom list into the Word document
$docx->addList($myList, 'latin');

$docx->addText('If the same list style is use in another list, the numbering continues:');

$docx->addList($myList, 'latin');

$docx->addText('A new custom list style must be generate to restart the numering:');

$docx->createListStyle('latin_2', $latinListOptions);
$docx->addList($myList, 'latin_2');

$docx->createDocx('example_createListStyle_7');