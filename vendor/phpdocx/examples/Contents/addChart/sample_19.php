<?php
// add a histogram chart

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->addText('Add a histogram chart to the Word document:');

$data = array(
    'data' => array(
        array(
            'values' => array(1, 3, 3, 3, 5, 6, 6, 6, 7, 8, 8, 9, 9, 9, 9, 9, 10, 10, 10, 10, 10, 10, 11, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13, 14, 14, 14, 14, 14, 14, 15, 15, 15, 15, 15, 15, 15, 15, 16, 16, 16, 16, 17, 17, 17, 17, 17, 17, 18, 18, 18, 18, 19, 19, 19, 20, 21, 22, 22, 24, 24),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'histogramChart',
    'title' => 'My custom title',
    'sizeX' => 15,
    'sizeY' => 8,
);
$docx->addChart($paramsChart);

$docx->addText('Add a histogram chart with a custom color and style to the Word document:');

$data = array(
    'legend' => array('Legend 1'),
    'data' => array(
        array(
            'values' => array(1, 3, 3, 3, 5, 6, 6, 6, 7, 8, 8, 9, 9, 9, 9, 9, 10, 10, 10, 10, 10, 10, 11, 11, 11, 11, 11, 11, 12, 12, 12, 12, 12, 12, 13, 13, 13, 13, 13, 14, 14, 14, 14, 14, 14, 15, 15, 15, 15, 15, 15, 15, 15, 16, 16, 16, 16, 17, 17, 17, 17, 17, 17, 18, 18, 18, 18, 19, 19, 19, 20, 21, 22, 22, 24, 24),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'histogramChart',
    'title' => 'My custom title',
    'color' => 'colorful3',
    'style' => 'style2',
    'chartAlign' => 'center',
    'sizeX' => 12,
    'sizeY' => 8,
    'legendPos' => 'b',
    'showLegend' => true,
);
$docx->addChart($paramsChart);

$docx->createDocx('example_addChart_19');