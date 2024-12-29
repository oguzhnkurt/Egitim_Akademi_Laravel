<?php

/**
 * Create histogram chart
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateHistogramChart extends CreateGraphicEx
{
    /**
     * Create XML chart
     *
     * @access public
     * @param mixed $rId
     * @param array $options
     */
    public function createChartEx($rId, $options)
    {
        $xmlChart = '';

        // chartData
        $xmlChart .= '<cx:chartData><cx:data id="0">';

        // values
        $xmlChart .= '<cx:numDim type="val">';
        $xmlChart .= '<cx:f>Sheet1!$B$2:$B$'.(count($options['data']['data'][0]['values'])+1).'</cx:f>';
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
        $xmlChart .= '<cx:chart><cx:title align="ctr" overlay="0" pos="t">'.$titleContent.'</cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="clusteredColumn"><cx:tx><cx:txData><cx:f>Sheet1!$A$1</cx:f><cx:v>'.$seriesName.'</cx:v></cx:txData></cx:tx><cx:dataId val="0"/><cx:layoutPr><cx:binning intervalClosed="r" /></cx:layoutPr></cx:series></cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="0" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis></cx:plotArea>'.$legendContent.'</cx:chart>';

        $xmlChart = sprintf($this->xmlChartSkeleton, $xmlChart);

        return $xmlChart;
    }
}