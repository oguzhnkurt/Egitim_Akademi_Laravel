<?php
// remove variables (placeholders) from an existing DOCX

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocxFromTemplate('../../files/TemplateVariables.docx');

// remove the whole paragraph that contains the 'FOOTERVAR' variable
$docx->removeTemplateVariable('FOOTERVAR', 'block', 'footer');

// remove the 'OTHERVAR' variable value
$docx->removeTemplateVariable('OTHERVAR', 'inline');

// remove the line containing the FOOTNOTEVAR variable in the footnote
$docx->removeTemplateVariable('FOOTNOTEVAR', 'block', 'footnote');

$docx->createDocx('example_removeTemplateVariable_1');