<?php

/**
 * Create Line Chart
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateLineChart extends CreateGraphic implements InterfaceGraphic
{
    /**
     * Construct
     *
     * @access public
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Create embedded xml chart
     *
     * @access public
     */
    public function createEmbeddedXmlChart()
    {
        $this->_xmlChart = '';
        $args = func_get_args();
        $this->generateCHARTSPACE();
        $this->generateDATE1904(1);
        $this->generateLANG();
        $color = 2;
        if ($this->_color) {
            $color = $this->_color;
        }
        $this->generateSTYLE($color);
        $this->generateCHART();
        if ($this->_title != '') {
            $this->generateTITLE();
            $this->generateTITLETX();
            $this->generateRICH();
            $this->generateBODYPR();
            $this->generateLSTSTYLE();
            $this->generateTITLEP();
            $this->generateTITLEPPR();
            $this->generateDEFRPR('title');
            $this->generateTITLER();
            $this->generateTITLERPR();
            $this->generateTITLET($this->_title);
            $this->generateTITLELAYOUT();
            $this->cleanTemplateFonts();
        } else {
            $this->generateAUTOTITLEDELETED();
            $title = '';
        }
        if (strpos($this->_type, '3D') !== false) {
            $this->generateVIEW3D();
            $rotX = 30;
            $rotY = 30;
            $perspective = 30;
            if ($this->_rotX != '') {
                $rotX = $this->_rotX;
            }
            if ($this->_rotY != '') {
                $rotY = $this->_rotY;
            }
            if ($this->_perspective != '') {
                $perspective = $this->_perspective;
            }
            $this->generateROTX($rotX);
            $this->generateROTY($rotY);
            $this->generateRANGAX($this->_rAngAx);
            $this->generatePERSPECTIVE($perspective);
        }
        if ($this->values == '') {
            PhpdocxLogger::logger('Data values are missing.', 'fatal');
        }
        $this->generatePLOTAREA();
        $this->generateLAYOUT();

        if (strpos($this->_type, '3D') !== false) {
            $this->generateLINE3DCHART();
        } else {
            $this->generateLINECHART();
        }
        switch ($this->_style) {
            case 'radar':
                $style = 'marker';
                $marker = false;
                break;
            case 'marker':
            case 'filled':
                $style = $this->_style;
                break;
            default:
                $style = 'marker';
                break;
        }
        $groupBar = 'standard';
        if (!empty($this->_groupBar) && in_array($this->_groupBar, array('stacked', 'standard', 'percentStacked'))) {
            $groupBar = $this->_groupBar;
        }
        $this->generateGROUPING($groupBar);
        if (isset($this->values['legend'])) {
            $legends = $this->values['legend'];
        } else {
            echo('You haven`t added legends');
            return false;
        }
        $numValues = count($this->values['data']);
        $this->generateVARYCOLORS($this->_varyColors);
        $letter = 'A';
        // keep the max id value to be used with combo charts
        $idxMax = 0;
        for ($i = 0; $i < count($legends); $i++) {
            $this->generateSER();
            $this->generateIDX($i);
            $this->generateORDER($i);
            $letter++;

            $this->generateTX();
            $this->generateSTRREF();
            $this->generateF('Sheet1!$' . $letter . '$1');
            $this->generateSTRCACHE();
            $this->generatePTCOUNT();
            $this->generatePT();
            $this->generateV($legends[$i]);

            if (!empty($this->_symbol)) {
                $sizeSimbol = NULL;
                if (is_numeric($this->_symbolSize)) {
                    $sizeSimbol = $this->_symbolSize;
                }
                if (is_array($this->_symbol)) {
                    $this->generateMARKER($this->_symbol[$i], $sizeSimbol);
                } else {
                    $this->generateMARKER($this->_symbol, $sizeSimbol);
                }
            }
            $this->cleanTemplate2();

            if (isset($this->values['trendline'][$i])) {
                $this->generateTRENDLINE($this->values['trendline'][$i]);
            }

            if (is_array($this->_theme) && isset($this->_theme['serRgbColors']) && isset($this->_theme['serRgbColors'][$i])) {
                if ($this->_theme['serRgbColors'][$i] != null) {
                    $this->generateSPPR_SER();
                    $this->generateSPPR_SOLIDFILL($this->_theme['serRgbColors'][$i]);
                }
            }

            if (is_array($this->_theme) && isset($this->_theme['valueRgbColors']) && isset($this->_theme['valueRgbColors'][$i]) && $this->_theme['valueRgbColors'][$i] != null) {
                if ($this->_theme['valueRgbColors'][$i] != null) {
                    $this->generateCDPT($this->_theme['valueRgbColors'][$i]);
                }
            }

            if (is_array($this->_theme) && isset($this->_theme['serDataLabels'])) {
                if ($this->_theme['serDataLabels'][$i] != null) {
                    $this->generateDATALABELS_SER($this->_theme['serDataLabels'][$i], $i);
                }
            }

            if (is_array($this->_theme) && isset($this->_theme['valueDataLabels'])) {
                if (isset($this->_theme['valueDataLabels'][$i]) && $this->_theme['valueDataLabels'][$i] != null) {
                    $this->generateDATALABELS_DLBL($this->_theme['valueDataLabels'][$i], $i);
                }
            }

            $this->generateCAT();
            $this->generateSTRREF();
            $this->generateF('Sheet1!$A$2:$A$' . ($numValues + 1));
            $this->generateSTRCACHE();
            $this->generatePTCOUNT($numValues);

            $num = 0;
            foreach ($this->values['data'] as $value) {
                $this->generatePT($num);
                $this->generateV($value['name']);
                $num++;
            }
            $this->cleanTemplate2();
            if ($this->_type == 'radar' && $style == 'marker' && $marker == false) {
                $this->generateMARKER();
            }
            $this->generateVAL();
            $this->generateNUMREF();
            $this->generateF('Sheet1!$' . $letter . '$2:$' . $letter . '$' . ($numValues + 1));
            $this->generateNUMCACHE();
            $this->generateFORMATCODE();
            $this->generatePTCOUNT($numValues);
            $num = 0;
            foreach ($this->values['data'] as $name => $value) {
                $this->generatePT($num);
                $this->generateV($value['values'][$i]);
                $num++;
            }

            if (!empty($this->_smooth) && $this->_smooth) {
                $this->generateSMOOTH();
            } else if ($this->_smooth === '0') {
                $this->generateSMOOTH(0);
            }
            $this->cleanTemplate3();

            $idxMax = $i;
        }

        //Generate labels
        $this->generateSERDLBLS();
        $this->generateSHOWLEGENDKEY($this->_showLegendKey);
        $this->generateSHOWVAL($this->_showValue);
        $this->generateSHOWCATNAME($this->_showCategory);
        $this->generateSHOWSERNAME($this->_showSeries);
        $this->generateSHOWPERCENT($this->_showPercent);
        $this->generateSHOWBUBBLESIZE($this->_showBubbleSize);

        if (isset($args[0][1]['groupBar']) && ($args[0][1]['groupBar'] == 'stacked' ||
                $args[0][1]['groupBar'] == 'percentStacked')
        ) {
            $this->generateOVERLAP();
        }
        $this->generateAXID();
        $this->generateAXID(59040512);

        if ($this->_comboChart !== null) {
            $this->addComboChart($idxMax + 1);
        }

        if (strpos($this->_type, '3D') !== false) {
            $this->generateAXID(83319040);
        }
        $this->generateVALAX();
        $this->generateAXAXID(59040512);
        $this->generateSCALING(true);
        $this->generateDELETE($this->_deleteAxisValues);
        if (!empty($this->_orientation) && is_array($this->_orientation) && isset($this->_orientation[0]) && !is_null($this->_orientation[0]))  {
            $this->generateORIENTATION($this->_orientation[0]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->_axPos) && is_array($this->_axPos) && isset($this->_axPos[0]) && !is_null($this->_axPos[0]))  {
            $this->generateAXPOS($this->_axPos[0]);
        } else {
            $this->generateAXPOS('l');
        }

        switch ($this->_hgrid) {
            case 1:
                $this->generateMAJORGRIDLINES();
                break;
            case 2:
                $this->generateMINORGRIDLINES();
                break;
            case 3:
                $this->generateMAJORGRIDLINES();
                $this->generateMINORGRIDLINES();
                break;
            default:
                break;
        }

        if (!empty($this->_vaxLabel)) {
            $this->generateAXLABEL($this->_vaxLabel);
            $vert = 'horz';
            $rot = 0;
            if ($this->_vaxLabelDisplay == 'vertical') {
                $vert = 'wordArtVert';
            }
            if ($this->_vaxLabelDisplay == 'rotated') {
                $rot = '-5400000';
            }
            $this->generateAXLABELDISP($vert, $rot);
        }
        if ($this->_formatCode) {
            $this->generateNUMFMT($this->_formatCode, 0);
        } else {
            $this->generateNUMFMT();
        }
        if (!is_array($this->_tickLblPos)) {
            $this->generateTICKLBLPOS($this->_tickLblPos, true);
        } else if (!empty($this->_tickLblPos) && is_array($this->_tickLblPos) && isset($this->_tickLblPos[0]) && !is_null($this->_tickLblPos[0])) {
            $this->generateTICKLBLPOS($this->_tickLblPos[0]);
        }
        $this->generateMAJORUNIT($this->_majorUnit);
        $this->generateMINORUNIT($this->_minorUnit);
        $this->generateCROSSAX(59034624);
        $this->generateCROSSES();
        $this->generateCROSSBETWEEN();
        $this->generateCATAX();
        $this->generateAXAXID(59034624);
        $this->generateSCALING();
        $this->generateDELETE($this->_deleteAxisValues);
        if (!empty($this->_orientation) && is_array($this->_orientation) && isset($this->_orientation[1]) && !is_null($this->_orientation[1]))  {
            $this->generateORIENTATION($this->_orientation[1]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->_axPos) && is_array($this->_axPos) && isset($this->_axPos[1]) && !is_null($this->_axPos[1]))  {
            $this->generateAXPOS($this->_axPos[1]);
        } else {
            $this->generateAXPOS();
        }

        if (!empty($this->_haxLabel)) {
            $this->generateAXLABEL($this->_haxLabel);
            $vert = 'horz';
            $rot = 0;
            if ($this->_haxLabelDisplay == 'vertical') {
                $vert = 'wordArtVert';
            }
            if ($this->_haxLabelDisplay == 'rotated') {
                $rot = '-5400000';
            }
            $this->generateAXLABELDISP($vert, $rot);
        }
        switch ($this->_vgrid) {
            case 1:
                $this->generateMAJORGRIDLINES();
                break;
            case 2:
                $this->generateMINORGRIDLINES();
                break;
            case 3:
                $this->generateMAJORGRIDLINES();
                $this->generateMINORGRIDLINES();
                break;
            default:
                break;
        }

        if (strpos($this->_type, 'surface') !== false) {
            $this->generateMAJORTICKMARK();
        }
        if (!is_array($this->_tickLblPos)) {
            $this->generateTICKLBLPOS();
        } else if (!empty($this->_tickLblPos) && is_array($this->_tickLblPos) && isset($this->_tickLblPos[1]) && !is_null($this->_tickLblPos[1])) {
            $this->generateTICKLBLPOS($this->_tickLblPos[1]);
        }
        $this->generateCROSSAX();
        $this->generateCROSSES();
        $this->generateAUTO();
        $this->generateLBLALGN();
        $this->generateLBLOFFSET();
        if (!empty($this->_showTable)) {
            $this->generateDATATABLE();
        }

        $this->generateLEGEND();
        $this->generateLEGENDPOS($this->_legendPos);
        $this->generateLEGENDOVERLAY($this->_legendOverlay);

        $this->generatePLOTVISONLY();

        if ((!isset($this->_border) || $this->_border == 0 || !is_numeric($this->_border))
        ) {
            $this->generateSPPR();
            $this->generateLN();
            $this->generateNOFILL();
        } else {
            $this->generateSPPR();
            $this->generateLN($this->_border);
        }

        if ($this->_font != '') {
            $this->generateTXPR();
            $this->generateLEGENDBODYPR();
            $this->generateLSTSTYLE();
            $this->generateAP();
            $this->generateAPPR();
            $this->generateDEFRPR();
            $this->generateRFONTS($this->_font);
            $this->generateENDPARARPR();
        }


        $this->generateEXTERNALDATA();
        $this->cleanTemplateDocument();
        return $this->_xmlChart;
    }

    public function dataTag()
    {
        return array('val');
    }

    /**
     * Return the type of the xlsx object
     *
     * @access public
     */
    public function getXlsxType()
    {
        return CreateCompletedXlsx::getInstance();
    }

}
