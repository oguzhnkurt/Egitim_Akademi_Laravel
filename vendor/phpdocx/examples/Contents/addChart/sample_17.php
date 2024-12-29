<?php
// add a funnel chart

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->addText('Add a funnel chart to the Word document:');

$data = array(
    'data' => array(
        array(
            'name' => 'Category 1',
            'values' => array(3000),
        ),
        array(
            'name' => 'Category 2',
            'values' => array(4000),
        ),
        array(
            'name' => 'Category 3',
            'values' => array(4500),
        ),
        array(
            'name' => 'Category 4',
            'values' => array(1200),
        ),
        array(
            'name' => 'Category 5',
            'values' => array(250),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'funnelChart',
    'title' => 'My custom title',
    'sizeX' => 15,
    'sizeY' => 8,
);
$docx->addChart($paramsChart);

$docx->addText('Add a funnel chart with a custom color and style to the Word document:');

$data = array(
    'legend' => array('Legend 1'),
    'data' => array(
        array(
            'name' => 'Category A',
            'values' => array(3000),
        ),
        array(
            'name' => 'Category B',
            'values' => array(4000),
        ),
        array(
            'name' => 'Category C',
            'values' => array(4500),
        ),
        array(
            'name' => 'Category D',
            'values' => array(1200),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'funnelChart',
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

$docx->createDocx('example_addChart_17');