<?php
// add structured document tags

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

// Combo box
$list = array(
	array('First Choice', 1),
	array('second choice', 2),
	array('Third choice', 3),
);
$options = array(
	'listItems' => $list,
	'placeholderText' => 'Choose a value or write it down',
	'alias' => 'Combo Box',
	'fontSize' => 12,
	'italic' => true,
	'color' => 'FF0000',
	'bold' => true,
	'underline' => 'single',
	'font' => 'Algerian'
);
$docx->addStructuredDocumentTag('comboBox', $options );

// date
$options = array(
	'placeholderText' => 'Choose a date',
	'alias' => 'Date picker',
	'fontSize' => 14,
	'italic' => true,
	'color' => '777777',
	'bold' => true,
	'font' => 'Calibri'
);
$docx->addStructuredDocumentTag('date', $options);

// dropdown
$list = array(
	array('One', 1),
	array('Two', 2),
	array('Three', 3)
);
$options = array(
	'listItems' => $list,
	'placeholderText' => 'Choose a value',
	'alias' => 'Dropdown menu',
	'fontSize' => 12
);
$docx->addStructuredDocumentTag('comboBox', $options);

// richText
$options = array(
	'placeholderText' => 'This text is locked',
	'alias' => 'Rich text',
	'lock' => 'contentLocked'
);
$docx->addStructuredDocumentTag('richText', $options);

// checkboxes with custom styles
$docx->addStructuredDocumentTag('checkbox', array('fontSize' => 18, 'checked' => true));
$docx->addStructuredDocumentTag('checkbox', array('fontSize' => 18, 'checked' => true, 'highlightColor' => 'red'));
// checkbox with custom styles and custom symbol
$docx->addStructuredDocumentTag('checkbox', array('fontSize' => 18, 'checked' => true, 'checkedState' => array('font' => 'Wingdings', 'value' => '00FE'), 'uncheckedState' => array('font' => 'Wingdings', 'value' => '006F'), 'sym' => array('char' => '00FE', 'font' => 'Wingdings')));

// checkbox added as WordFragment
$checkboxFragment = new WordFragment($docx, 'document');
$checkboxFragment->addStructuredDocumentTag('checkbox', ['fontSize' => 11, 'checked' => true]);
$docx->addText([array('text' => 'Checkbox: ', 'fontSize' => 12), $checkboxFragment]);

$docx->createDocx('example_addStructuredDocumentTag_1');