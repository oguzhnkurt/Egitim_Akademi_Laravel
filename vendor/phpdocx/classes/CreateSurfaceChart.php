<?php

/**
 * Create Surface Chart
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateSurfaceChart extends CreateGraphic implements InterfaceGraphic
{
    /**
     * Create embedded xml chart
     *
     * @access public
     */
    public function createEmbeddedXmlChart()
    {
        $this->_xmlChart = '';
        $this->generateCHARTSPACE();
        $this->generateDATE1904(1);
        $this->generateLANG();
        $this->generateROUNDEDCORNERS($this->_roundedCorners);
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
        } else {
            $this->generateAUTOTITLEDELETED();
            $title = '';
        }
        $this->generateVIEW3D();
        if (strpos($this->_type, '3D') !== false) {
            $this->generatePERSPECTIVE(30);
        } else {
            $this->generateROTX(90);
            $this->generateROTY(0);
            $this->generateRANGAX($this->_rAngAx);
            $this->generatePERSPECTIVE(0);
        }

        if ($this->values == '') {
            PhpdocxLogger::logger('Data values are missing.', 'fatal');
        }
        $this->generatePLOTAREA();
        $this->generateLAYOUT();

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

        if (strpos($this->_type, '3D') !== false) {
            $this->generateSURFACE3DCHART();
        } else {
            $this->generateSURFACECHART();
        }
        if ($this->_wireframe)
            $this->generateWIREFRAME();
        if (isset($this->values['legend'])) {
            $legends = $this->values['legend'];
        } else {
            echo('You haven`t added legends');
            return false;
        }
        $numValues = count($this->values['data']);
        $this->generateVARYCOLORS($this->_varyColors);
        $letter = 'A';
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

            $this->cleanTemplate2();

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

            $this->cleanTemplate3();
        }

        //Generate labels
        $this->generateSERDLBLS();
        $this->generateSHOWLEGENDKEY($this->_showLegendKey);
        $this->generateSHOWVAL($this->_showValue);
        $this->generateSHOWCATNAME($this->_showCategory);
        $this->generateSHOWSERNAME($this->_showSeries);
        $this->generateSHOWPERCENT($this->_showPercent);
        $this->generateSHOWBUBBLESIZE($this->_showBubbleSize);

        $this->generateBANDFMTS();
        $this->generateAXID();
        $this->generateAXID(59040512);
        $this->generateAXID(83319040);
        $this->generateVALAX();
        $this->generateAXAXID(59040512);
        $this->generateSCALING();
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
        $this->generateMAJORTICKMARK('none');
        if (!is_array($this->_tickLblPos)) {
            $this->generateTICKLBLPOS('none');
        } else if (!empty($this->_tickLblPos) && is_array($this->_tickLblPos) && isset($this->_tickLblPos[0]) && !is_null($this->_tickLblPos[0])) {
            $this->generateTICKLBLPOS($this->_tickLblPos[0]);
        }
        $this->generateCROSSAX(59034624);
        $this->generateCROSSES();
        $this->generateCROSSBETWEEN('midCat');
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

        $this->generateMAJORTICKMARK();
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

        $this->generateSERAX();
        $this->generateAXAXID(83319040);
        $this->generateSCALING();
        $this->generateDELETE($this->_deleteAxisValues);
        if (!empty($this->_orientation) && is_array($this->_orientation) && isset($this->_orientation[2]) && !is_null($this->_orientation[2]))  {
            $this->generateORIENTATION($this->_orientation[2]);
        } else {
            $this->generateORIENTATION();
        }
        if (!empty($this->_axPos) && is_array($this->_axPos) && isset($this->_axPos[2]) && !is_null($this->_axPos[2]))  {
            $this->generateAXPOS($this->_axPos[2]);
        } else {
            $this->generateAXPOS('b');
        }
        switch ($this->_hgrid) {
            case 2:
                $this->generateMINORGRIDLINES();
                break;
            case 3:
                $this->generateMAJORGRIDLINES();
                $this->generateMINORGRIDLINES();
                break;
            case 1:
            default:
                $this->generateMAJORGRIDLINES();
                break;
        }
        $this->generateMAJORTICKMARK();
        if (!is_array($this->_tickLblPos)) {
            $this->generateTICKLBLPOS();
        } else if (!empty($this->_tickLblPos) && is_array($this->_tickLblPos) && isset($this->_tickLblPos[2]) && !is_null($this->_tickLblPos[2])) {
            $this->generateTICKLBLPOS($this->_tickLblPos[2]);
        }
        $this->generateCROSSAX(59034624);
        $this->generateCROSSES();
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
     * retrun the type of the xlsx object
     *
     * @access public
     */
    public function getXlsxType()
    {
        return CreateCompletedXlsx::getInstance();
    }

}
