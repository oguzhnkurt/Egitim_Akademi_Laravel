<?php
// add a box and whisker chart

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->addText('Add a box and whisker chart to the Word document:');

$data = array(
    'data' => array(
        array(
            array(
                'name' => 'Category 1',
                'values' => array(-7, -10, -28, 47, 11, -24, -24, 36, 10),
            ),
            array(
                'name' => 'Category 2',
                'values' => array(-78, 47, -24, -17, -12, -11, 17),
            ),
            array(
                'name' => 'Category 3',
                'values' => array(14, 46, -18, 19, -26, -20),
            ),
        ),
        array(
            array(
                'name' => 'Category 1',
                'values' => array(-3, 1, -6, 10, 34, 128, 22, -12, -28),
            ),
            array(
                'name' => 'Category 2',
                'values' => array(6, 31, 3, 12, -12, -13, 6),
            ),
            array(
                'name' => 'Category 3',
                'values' => array(15, 41, 16, 10, 23, 16),
            ),
        ),
        array(
            array(
                'name' => 'Category 1',
                'values' => array(-24, 11, 34, -19, 4, 27, 27, -3, 44),
            ),
            array(
                'name' => 'Category 2',
                'values' => array(50, 91, -8, 36, 16, 24, 46),
            ),
            array(
                'name' => 'Category 3',
                'values' => array(14, -6, 48, 23, 23, -18),
            ),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'boxWhiskerChart',
    'title' => 'My custom title',
    'sizeX' => 15,
    'sizeY' => 8,
);
$docx->addChart($paramsChart);

$docx->addText('Add a box and whisker chart with a custom color and style to the Word document:');

$data = array(
    'legend' => array('Legend 1', 'Legend 2', 'Legend 3'),
    'data' => array(
        array(
            array(
                'name' => 'Category 1',
                'values' => array(-7, -10, -28, 47, 11, -24, -24, 36, 10),
            ),
            array(
                'name' => 'Category 2',
                'values' => array(-78, 47, -24, -17, -12, -11, 17),
            ),
            array(
                'name' => 'Category 3',
                'values' => array(14, 46, -18, 19, -26, -20),
            ),
        ),
        array(
            array(
                'name' => 'Category 1',
                'values' => array(-3, 1, -6, 10, 34, 128, 22, -12, -28),
            ),
            array(
                'name' => 'Category 2',
                'values' => array(6, 31, 3, 12, -12, -13, 6),
            ),
            array(
                'name' => 'Category 3',
                'values' => array(15, 41, 16, 10, 23, 16),
            ),
        ),
        array(
            array(
                'name' => 'Category 1',
                'values' => array(-24, 11, 34, -19, 4, 27, 27, -3, 44),
            ),
            array(
                'name' => 'Category 2',
                'values' => array(50, 91, -8, 36, 16, 24, 46),
            ),
            array(
                'name' => 'Category 3',
                'values' => array(14, -6, 48, 23, 23, -18),
            ),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'boxWhiskerChart',
    'title' => 'A new title',
    'color' => 'colorful2',
    'style' => 'style10',
    'chartAlign' => 'center',
    'sizeX' => 12,
    'sizeY' => 8,
    'legendPos' => 'b',
    'showLegend' => true,
);
$docx->addChart($paramsChart);

$docx->createDocx('example_addChart_20');