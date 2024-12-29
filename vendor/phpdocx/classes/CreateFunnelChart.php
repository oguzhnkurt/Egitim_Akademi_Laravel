<?php

/**
 * Create funnel chart
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateFunnelChart extends CreateGraphicEx
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
        $xmlChart .= '<cx:f>Sheet1!$A$2:$A$'.(count($options['data']['data'])+1).'</cx:f>';
        $xmlChart .= '<cx:lvl ptCount="'.count($options['data']['data']).'">';
        $idx = 0;
        foreach ($options['data']['data'] as $values) {
            $xmlChart .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($values['name']).'</cx:pt>';
            $idx++;
        }
        $xmlChart .= '</cx:lvl></cx:strDim>';
        // values
        $xmlChart .= '<cx:numDim type="val">';
        $xmlChart .= '<cx:f>Sheet1!$B$2:$B$'.(count($options['data']['data'])+1).'</cx:f>';
        $xmlChart .= '<cx:lvl ptCount="'.count($options['data']['data']).'">';
        $idx = 0;
        foreach ($options['data']['data'] as $values) {
            $xmlChart .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($values['values'][0]).'</cx:pt>';
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
        $xmlChart .= '<cx:chart><cx:title align="ctr" overlay="0" pos="t">'.$titleContent.'</cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="funnel"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>'.$seriesName.'</cx:v></cx:txData></cx:tx><cx:dataLabels><cx:visibility categoryName="0" seriesName="0" value="1"/></cx:dataLabels><cx:dataId val="0"/></cx:series></cx:plotAreaRegion><cx:axis id="1"><cx:catScaling gapWidth="0.0599999987"/><cx:tickLabels/></cx:axis></cx:plotArea>'.$legendContent.'</cx:chart>';

        $xmlChart = sprintf($this->xmlChartSkeleton, $xmlChart);

        return $xmlChart;
    }
}