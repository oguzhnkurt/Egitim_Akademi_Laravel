<?php
// convert HTML to DOCX using CSS variables

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$html = '
<style>
:root {
    --blue: #1990fa;
    font-weight: bold;
}
.varcolor {
    color: var(--blue);
}
.color {
    color: green;
}
</style>
<p class="varcolor">Text with blue color using a CSS variable.</p>
<p class="color">Text with a custom color.</p>
';
$docx->embedHTML($html, array('parseCSSVars' => true));

$html = '
<style>
:root {
    --text-color: green;
    --myborders: dashed;
    --text-size: 4px;
    font-weight: bold;
    --color-r: red;
    --color-g: green;
    --sz-t1: 46px;
    --font-fc: Arial;
}

p {
    --background-cl: #decb97;
    border-bottom: var(--text-size) var(--myborders) var(--text-color);
    color: var(--color-r);
    font-size: var(--sz-t1, 32px);
    background-color: var(--background-cl);
}

span.c1 {
    color: var(--color-g);
    font-size: var(--sz-t2, 16px);
    font-family: var(--font-fc);
}

</style>
<p>Complex <span class="c1">CSS variables</span> can be used.</p>
';
$docx->embedHTML($html, array('parseCSSVars' => true));

$docx->createDocx('example_embedHTML_5');