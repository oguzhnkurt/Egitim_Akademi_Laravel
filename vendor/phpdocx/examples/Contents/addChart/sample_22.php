<?php
// add a treemap chart

require_once '../../../classes/CreateDocx.php';

$docx = new CreateDocx();

$docx->addText('Add a treemap chart to the Word document:');

$data = array(
    'data' => array(
        array(
            'leaf' => array('Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', 'Leaf 6', 'Leaf 7', 'Leaf 8', 'Leaf 9', 'Leaf 10', 'Leaf 11', 'Leaf 12', 'Leaf 13', 'Leaf 14', 'Leaf 15', 'Leaf 16'),
            'stem' => array('Stem 1', 'Stem 1', 'Stem 1', 'Stem 2', 'Stem 2', 'Stem 2', 'Stem 2', 'Stem 3', 'Stem 3', 'Stem 4', 'Stem 4', 'Stem 5', 'Stem 5', 'Stem 6', 'Stem 6', 'Stem 6'),
            'branch' => array('Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 3', 'Branch 3', 'Branch 3', 'Branch 3', 'Branch 3'),
            'values' => array(22, 12, 18, 87, 88, 17, 9, 25, 23, 24, 89, 16, 19, 86, 10, 11),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'treemapChart',
    'title' => 'My custom title',
    'style' => 'style11',
    'sizeX' => 15,
    'sizeY' => 8,
);
$docx->addChart($paramsChart);

$docx->addText('Add a treemap chart with a custom color and style to the Word document:');

$data = array(
    'data' => array(
        array(
            'leaf' => array('Leaf 1', 'Leaf 2', 'Leaf 3', 'Leaf 4', 'Leaf 5', 'Leaf 6', 'Leaf 7', 'Leaf 8', 'Leaf 9', 'Leaf 10', 'Leaf 11', 'Leaf 12', 'Leaf 13', 'Leaf 14', 'Leaf 15', 'Leaf 16'),
            'stem' => array('Stem 1', 'Stem 1', 'Stem 1', 'Stem 2', 'Stem 2', 'Stem 2', 'Stem 2', 'Stem 3', 'Stem 3', 'Stem 4', 'Stem 4', 'Stem 5', 'Stem 5', 'Stem 6', 'Stem 6', 'Stem 6'),
            'branch' => array('Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 1', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 2', 'Branch 3', 'Branch 3', 'Branch 3', 'Branch 3', 'Branch 3'),
            'values' => array(22, 12, 18, 87, 88, 17, 9, 25, 23, 24, 89, 16, 19, 86, 10, 11),
        ),
    ),
);

$paramsChart = array(
    'data' => $data,
    'type' => 'treemapChart',
    'title' => 'A new title',
    'color' => 'monochromatic3',
    'style' => 'style11',
    'chartAlign' => 'center',
    'sizeX' => 12,
    'sizeY' => 8,
    'legendPos' => 'b',
    'showLegend' => true,
);
$docx->addChart($paramsChart);

$docx->createDocx('example_addChart_22');