<?php
// add a text applying styles, using parseLineBreaks to parse line breaks and parseTabs to parse tabs

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();
$text = 'Lorem ipsum dolor sit amet, \nconsectetur adipisicing elit, \n' .
    'sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut ' .
    'enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut' .
    'aliquip ex ea commodo consequat. \n\Duis aute irure dolor in reprehenderit ' .
    'in voluptate velit esse cillum dolore eu fugiat nulla pariatur. ' .
    'Excepteur sint occaecat cupidatat non proident, sunt in culpa qui ' .
    'officia deserunt mollit anim id est laborum.';

$paragraphOptions = array(
    'bold' => true,
    'font' => 'Arial',
    'parseLineBreaks' => true,
);

$docx->addText($text, $paragraphOptions);

$text = "Lorem \tipsum dolor sit\n\tamet.";

$paragraphOptions = array(
    'bold' => true,
    'font' => 'Arial',
    'parseTabs' => true,
    'parseLineBreaks' => true,
);

$docx->addText($text, $paragraphOptions);

$docx->createDocx('example_addText_6');