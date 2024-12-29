<?php

/**
 * Create treemap chart
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateTreemapChart extends CreateGraphicEx
{
    /**
     * Create XML chart
     *
     * @access public
     * @param mixed $rId
     * @param array $options
     * @return string
     */
    public function createChartEx($rId, $options)
    {
        $xmlChart = '';

        // chartData
        $xmlChart .= '<cx:chartData><cx:data id="0">';

        // categories
        $xmlChart .= '<cx:strDim type="cat">';
        $xmlChart .= '<cx:f>Sheet1!$A$2:$C$'.(count($options['data']['data'][0]['values'])+1).'</cx:f>';
        // leaf
        if (isset($options['data']['data'][0]['leaf'])) {
            $xmlChart .= '<cx:lvl ptCount="'.count($options['data']['data'][0]['leaf']).'">';
            $idx = 0;
            foreach ($options['data']['data'][0]['leaf'] as $value) {
                $xmlChart .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($value).'</cx:pt>';
                $idx++;
            }
            $xmlChart .= '</cx:lvl>';
        }
        // stem
        if (isset($options['data']['data'][0]['stem'])) {
            $xmlChart .= '<cx:lvl ptCount="'.count($options['data']['data'][0]['stem']).'">';
            $idx = 0;
            foreach ($options['data']['data'][0]['stem'] as $value) {
                $xmlChart .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($value).'</cx:pt>';
                $idx++;
            }
            $xmlChart .= '</cx:lvl>';
        }
        // branch
        if (isset($options['data']['data'][0]['branch'])) {
            $xmlChart .= '<cx:lvl ptCount="'.count($options['data']['data'][0]['branch']).'">';
            $idx = 0;
            foreach ($options['data']['data'][0]['branch'] as $value) {
                $xmlChart .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($value).'</cx:pt>';
                $idx++;
            }
            $xmlChart .= '</cx:lvl>';
        }
        $xmlChart .= '</cx:strDim>';
        // values
        $xmlChart .= '<cx:numDim type="size">';
        $xmlChart .= '<cx:f>Sheet1!$D$2:$D$'.(count($options['data']['data'][0]['values'])+1).'</cx:f>';
        $xmlChart .= '<cx:lvl ptCount="'.count($options['data']['data'][0]['values']).'">';
        $idx = 0;
        foreach ($options['data']['data'][0]['values'] as $value) {
            $xmlChart .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($value).'</cx:pt>';
            $idx++;
        }
        $xmlChart .= '</cx:lvl></cx:numDim>';

        $xmlChart .= '</cx:data></cx:chartData>';

        // title
        $titleContent = '';
        if (isset($options['title'])) {
            $titleContent = $this->returnTitleContent($options['title']);
        }

        // legend
        $seriesName = 'Series1';
        if (isset($options['data']['legend']) && isset($options['data']['legend'][0])) {
            $seriesName = $this->parseAndCleanTextString($options['data']['legend'][0]);
        }
        // show legend
        $legendContent = '';
        if (isset($options['showLegend']) && $options['showLegend']) {
            $legendContent = $this->returnShowLegendContent($options);
        }

        // chart
        $xmlChart .= '<cx:chart><cx:title align="ctr" overlay="0" pos="t">'.$titleContent.'</cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="treemap"><cx:tx><cx:txData><cx:f>Sheet1!$D$1</cx:f><cx:v>'.$seriesName.'</cx:v></cx:txData></cx:tx><cx:dataLabels pos="inEnd"><cx:visibility seriesName="0" categoryName="1" value="0" /></cx:dataLabels><cx:dataId val="0" /><cx:layoutPr><cx:parentLabelLayout val="overlapping" /></cx:layoutPr></cx:series></cx:plotAreaRegion></cx:plotArea>'.$legendContent.'</cx:chart>';

        $xmlChart = sprintf($this->xmlChartSkeleton, $xmlChart);

        return $xmlChart;
    }
}