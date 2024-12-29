<?php

/**
 * Create box & whisker chart
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateBoxWhiskerChart extends CreateGraphicEx
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
        $xmlChart .= '<cx:chartData>';

        $idData = 0;
        $valuesLetter = 'B';
        foreach ($options['data']['data'] as $dataCategoriesValues) {
            $xmlChart .= '<cx:data id="'.$idData.'">';
            $categoriesContent = '';
            $valuesContent = '';
            $idx = 0;
            $ptCount = 0;

            foreach ($dataCategoriesValues as $dataCategoryValue) {
                foreach ($dataCategoryValue['values'] as $value) {
                    $categoriesContent .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($dataCategoryValue['name']).'</cx:pt>';
                    $valuesContent .= '<cx:pt idx="'.$idx.'">'.$this->parseAndCleanTextString($value).'</cx:pt>';

                    $ptCount++;
                    $idx++;
                }
            }

            // categories
            $xmlChart .= '<cx:strDim type="cat">';
            $xmlChart .= '<cx:f>Sheet1!$A$2:$A$'.($ptCount+1).'</cx:f>';
            $xmlChart .= '<cx:lvl ptCount="'.$ptCount.'">';
            $xmlChart .= $categoriesContent;
            $xmlChart .= '</cx:lvl></cx:strDim>';
            // values
            $xmlChart .= '<cx:numDim type="val">';
            $xmlChart .= '<cx:f>Sheet1!$'.$valuesLetter.'$2:$'.$valuesLetter.'$'.($ptCount+1).'</cx:f>';
            $xmlChart .= '<cx:lvl ptCount="'.$ptCount.'">';
            $xmlChart .= $valuesContent;
            $xmlChart .= '</cx:lvl></cx:numDim>';

            $xmlChart .= '</cx:data>';

            $idData++;
            $valuesLetter++;
        }

        $xmlChart .= '</cx:chartData>';

        // title
        $titleContent = '';
        if (isset($options['title'])) {
            $titleContent = $this->returnTitleContent($options['title']);
        }

        // series
        $seriesContent = '';
        $indexSerie = 0;
        $seriesLetter = 'B';
        foreach ($options['data']['data'] as $dataCategoriesValues) {
            $seriesName = 'Series'.($indexSerie+1);
            if (isset($options['data']['legend'][$indexSerie])) {
                $seriesName = $options['data']['legend'][$indexSerie];
            }

            $seriesName = 'Series'.($indexSerie+1);
            if (isset($options['data']['legend']) && isset($options['data']['legend'][$indexSerie])) {
                $seriesName = $this->parseAndCleanTextString($options['data']['legend'][$indexSerie]);
            }

            $seriesContent .= '<cx:series layoutId="boxWhisker"><cx:tx><cx:txData><cx:f>Sheet1!$'.$seriesLetter.'$1</cx:f><cx:v>'.$seriesName.'</cx:v></cx:txData></cx:tx><cx:dataId val="'.$indexSerie.'" /><cx:layoutPr><cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1" /><cx:statistics quartileMethod="exclusive" /></cx:layoutPr></cx:series>';

            $indexSerie++;
            $seriesLetter++;
        }

        // show legend
        $legendContent = '';
        if (isset($options['showLegend']) && $options['showLegend']) {
            $legendContent = $this->returnShowLegendContent($options);
        }

        // chart
        $xmlChart .= '<cx:chart><cx:title align="ctr" overlay="0" pos="t">'.$titleContent.'</cx:title><cx:plotArea><cx:plotAreaRegion>'.$seriesContent.'</cx:plotAreaRegion><cx:axis id="0"><cx:catScaling gapWidth="1" /><cx:tickLabels /></cx:axis><cx:axis id="1"><cx:valScaling /><cx:majorGridlines /><cx:tickLabels /></cx:axis></cx:plotArea>'.$legendContent.'</cx:chart>';

        $xmlChart = sprintf($this->xmlChartSkeleton, $xmlChart);

        return $xmlChart;
    }
}