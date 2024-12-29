<?php
// add a waterfall chart

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->addText('Add a waterfall chart to the Word document:');

$data = array(
    'data' => array(
        array(
            'name' => 'Category 1',
            'values' => array(100),
        ),
        array(
            'name' => 'Category 2',
            'values' => array(20),
        ),
        array(
            'name' => 'Category 3',
            'values' => array(50),
        ),
        array(
            'name' => 'Category 4',
            'values' => array(-40),
        ),
        array(
            'name' => 'Category 5',
            'values' => array(130),
        ),
        array(
            'name' => 'Category 6',
            'values' => array(-60),
        ),
        array(
            'name' => 'Category 7',
            'values' => array(70),
        ),
        array(
            'name' => 'Category 8',
            'values' => array(140),
        ),
    ),
    'subtotals' => array(0, 4, 7),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'waterfallChart',
    'title' => 'My custom title',
    'sizeX' => 15,
    'sizeY' => 8,
);
$docx->addChart($paramsChart);

$docx->addText('Add a waterfall chart with a custom color and style to the Word document:');

$data = array(
    'legend' => array('Legend 1'),
    'data' => array(
        array(
            'name' => 'Category A',
            'values' => array(100),
        ),
        array(
            'name' => 'Category B',
            'values' => array(20),
        ),
        array(
            'name' => 'Category C',
            'values' => array(50),
        ),
        array(
            'name' => 'Category D',
            'values' => array(-40),
        ),
        array(
            'name' => 'Category E',
            'values' => array(130),
        ),
        array(
            'name' => 'Category F',
            'values' => array(-60),
        ),
        array(
            'name' => 'Category G',
            'values' => array(70),
        ),
        array(
            'name' => 'Category H',
            'values' => array(140),
        ),
    ),
    'subtotals' => array(0, 4, 7),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'waterfallChart',
    'title' => 'A new title',
    'color' => 'monochromatic3',
    'style' => 'style8',
    'chartAlign' => 'center',
    'sizeX' => 12,
    'sizeY' => 8,
    'legendPos' => 'b',
    'showLegend' => true,
);
$docx->addChart($paramsChart);

$docx->createDocx('example_addChart_18');