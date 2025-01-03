<?php

/**
 * Create tables
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateTable extends CreateElement
{
    /**
     * @access private
     * @var array
     * @static
     */
    private static $_borders = array('top', 'left', 'bottom', 'right');

    /**
     * @access private
     * @var mixed
     * @static
     */
    private static $_instance = NULL;

    /**
     * Construct
     *
     * @access public
     */
    public function __construct()
    {

    }

    /**
     * Destruct
     *
     * @access public
     */
    public function __destruct()
    {

    }

    /**
     *
     * @access public
     * @return string
     */
    public function __toString()
    {
        $this->cleanTemplate();
        return $this->_xml;
    }

    /**
     *
     * @access public
     * @return CreateTable
     * @static
     */
    public static function getInstance()
    {
        if (self::$_instance == NULL) {
            self::$_instance = new CreateTable();
        }
        return self::$_instance;
    }

    /**
     * Create table
     *
     * @access public
     */
    public function createTable()
    {
        $this->_xml = '';
        $args = func_get_args();
        $args[1] = CreateDocx::translateTableOptions2StandardFormat($args[1]);

        if (is_array($args[0])) {
            //Normalize table data
            $tableData = $this->parseTableData($args[0]);
            $this->generateTBL();
            $this->generateTBLPR();
            if (isset($args[1]['TBLSTYLEval'])) {
                $this->generateTBLSTYLE($args[1]['TBLSTYLEval']);
            } else {
                $this->generateTBLSTYLE();
            }
            if (isset($args[1]['float'])) {
                $this->generateTBLFLOAT($args[1]['float']);
            }
            $this->generateTBLOVERLAP();
            if (isset($args[1]['bidi'])) {
                $this->generateTBLBIDI($args[1]['bidi']);
            }
            if (isset($args[1]['tableWidth']) && is_array($args[1]['tableWidth'])) {
                $this->generateTBLW($args[1]['tableWidth']['type'], $args[1]['tableWidth']['value']);
            } else {
                $this->generateTBLW('pct', 100);
            }
            if (isset($args[1]['jc'])) {
                $this->generateJC($args[1]['jc']);
            }
            if (isset($args[1]['cellSpacing'])) {
                $this->generateTBLCELLSPACING($args[1]['cellSpacing']);
            }
            if (isset($args[1]['indent'])) {
                $this->generateTBLINDENT($args[1]['indent']);
            }
            if (isset($args[1]['border']) ||
                    isset($args[1]['border_sz']) ||
                    isset($args[1]['border_color']) ||
                    isset($args[1]['border_spacing'])) {
                $this->generateTBLBORDERS();
                $sz = 6;
                if (isset($args[1]['border_sz'])) {
                    $sz = $args[1]['border_sz'];
                }
                $color = 'auto';
                if (isset($args[1]['border_color'])) {
                    $color = $args[1]['border_color'];
                }
                $spacing = '0';
                if (isset($args[1]['border_spacing'])) {
                    $spacing = $args[1]['border_spacing'];
                }
                if (isset($args[1]['border'])) {
                    $border = $args[1]['border'];
                } else {
                    $border = 'solid';
                }
                if (!isset($args[1]['border_settings']) ||
                        $args[1]['border_settings'] == 'all' ||
                        $args[1]['border_settings'] == 'outside') {
                    $this->generateTBLTOP($border, $sz, $spacing, $color);
                    $this->generateTBLLEFT($border, $sz, $spacing, $color);
                    $this->generateTBLBOTTOM($border, $sz, $spacing, $color);
                    $this->generateTBLRIGHT($border, $sz, $spacing, $color);
                }
                if (!isset($args[1]['border_settings']) ||
                        $args[1]['border_settings'] == 'all' ||
                        $args[1]['border_settings'] == 'inside') {
                    $this->generateTBLINSIDEH($border, $sz, $spacing, $color);
                    $this->generateTBLINSIDEV($border, $sz, $spacing, $color);
                }
            }
            if (isset($args[1]['tableLayout'])) {
                $this->generateTBLLAYOUT($args[1]['tableLayout']);
            }
            if (isset($args[1]['cellMargin'])) {
                $this->generateTBLCELLMARGIN($args[1]['cellMargin']);
            }
            if (isset($args[1]['conditionalFormatting']) && is_array($args[1]['conditionalFormatting'])) {
                $this->generateTBLLOOK($args[1]['conditionalFormatting']);
            }
            if (isset($args[1]['descr'])) {
                $this->generateTBLDESCR($args[1]['descr']);
            }
            if (isset($args[1]['descrTitle'])) {
                $this->generateTBLDESCRTITLE($args[1]['descrTitle']);
            }


            $rowNumber = 0;
            $colNumber = 0;
            $this->generateTBLGRID();
            if (isset($args[1]['size_col']) && is_array($args[1]['size_col'])) {
                foreach ($args[1]['size_col'] as $key => $widthCol) {
                    $this->generateGRIDCOL($widthCol);
                }
            } else {
                foreach ($tableData as $row) {
                    $rowLength = array();
                    $rowLength[] = count($row);
                }
                if (isset($rowLength)) {
                    $numCols = max($rowLength);
                    for ($k = 0; $k < $numCols; $k++) {
                        if (isset($args[1]['size_col'])) {
                            $this->generateGRIDCOL($args[1]['size_col']);
                        } else {
                            $this->generateGRIDCOL();
                        }
                    }
                }
            }

            foreach ($tableData as $row) {
                $this->generateTR();
                $this->generateTRPR();
                if (isset($args[1]['cantSplitRows']) && $args[1]['cantSplitRows']) {
                    if (isset($args[2][$rowNumber]['cantSplit'])) {
                        $this->generateTRCANTSPLIT($args[2][$rowNumber]['cantSplit']);
                    } else {
                        $this->generateTRCANTSPLIT();
                    }
                } else if (isset($args[2][$rowNumber]['cantSplit']) &&
                        $args[2][$rowNumber]['cantSplit']) {
                    $this->generateTRCANTSPLIT();
                }
                if (isset($args[2][$rowNumber]['height']) ||
                        isset($args[2][$rowNumber]['minHeight'])) {
                    if (isset($args[2][$rowNumber]['height'])) {
                        $height = $args[2][$rowNumber]['height'];
                        $hRule = 'exact';
                    } else {
                        $height = $args[2][$rowNumber]['minHeight'];
                        $hRule = 'atLeast';
                    }
                    $this->generateTRHEIGHT($height, $hRule);
                }
                if (isset($args[2][$rowNumber]['tableHeader']) &&
                        $args[2][$rowNumber]['tableHeader']) {
                    $this->generateTRTABLEHEADER();
                }
                $rowNumber++;
                foreach ($row as $cellContent) {
                    $cellContent = CreateDocx::translateTableOptions2StandardFormat($cellContent);
                    $this->cleanTemplateTrPr();
                    $this->generateTC();
                    $this->generateTCPR();
                    if ($rowNumber == 1 && isset($args[1]['size_col'])) {
                        if (is_array($args[1]['size_col']) &&
                                isset($args[1]['size_col'][$colNumber])) {
                            $this->generateTCW($args[1]['size_col'][$colNumber]);
                        } else if (!is_array($args[1]['size_col'])) {
                            $this->generateTCW($args[1]['size_col']);
                        }
                    } else {
                        if (isset($cellContent['width']) && is_int($cellContent['width'])) {
                            $this->generateCELLWIDTH($cellContent['width']);
                        }
                    }
                    if (isset($cellContent['colspan']) && $cellContent['colspan'] > 1) {
                        $this->generateCELLGRIDSPAN($cellContent['colspan']);
                    }
                    if (isset($cellContent['rowspan']) && $cellContent['rowspan'] >= 1) {
                        if (isset($cellContent['vmerge']) && $cellContent['vmerge'] == 'continue') {
                            $this->generateCELLVMERGE('continue');
                        } else {
                            $this->generateCELLVMERGE('restart');
                        }
                    }
                    //we set the drawCellBorders to false
                    $drawCellBorders = false;
                    $border = array();

                    //Run over the general border properties
                    if (isset($cellContent['border'])) {
                        $drawCellBorders = true;
                        foreach (self::$_borders as $key => $value) {
                            $border[$value]['type'] = $cellContent['border'];
                        }
                    }
                    if (isset($cellContent['border_color'])) {
                        $drawCellBorders = true;
                        foreach (self::$_borders as $key => $value) {
                            $border[$value]['color'] = $cellContent['border_color'];
                        }
                    }
                    if (isset($cellContent['border_spacing'])) {
                        $drawCellBorders = true;
                        foreach (self::$_borders as $key => $value) {
                            $border[$value]['spacing'] = $cellContent['border_spacing'];
                        }
                    }
                    if (isset($cellContent['border_sz'])) {
                        $drawCellBorders = true;
                        foreach (self::$_borders as $key => $value) {
                            $border[$value]['sz'] = $cellContent['border_sz'];
                        }
                    }
                    //Run over the border choices of each side
                    foreach (self::$_borders as $key => $value) {
                        if (isset($cellContent['border_' . $value])) {
                            $drawCellBorders = true;
                            $border[$value]['type'] = $cellContent['border_' . $value];
                        }
                        if (isset($cellContent['border_' . $value . '_color'])) {
                            $drawCellBorders = true;
                            $border[$value]['color'] = $cellContent['border_' . $value . '_color'];
                        }
                        if (isset($cellContent['border_' . $value . '_spacing'])) {
                            $drawCellBorders = true;
                            $border[$value]['spacing'] = $cellContent['border_' . $value . '_spacing'];
                        }
                        if (isset($cellContent['border_' . $value . '_sz'])) {
                            $drawCellBorders = true;
                            $border[$value]['sz'] = $cellContent['border_' . $value . '_sz'];
                        }
                    }
                    if ($drawCellBorders) {
                        $this->generateCELLBORDER($border);
                    }
                    if (isset($cellContent['background_color'])) {
                        $this->generateCELLSHD($cellContent['background_color']);
                    }
                    if (isset($cellContent['noWrap'])) {
                        $this->generateCELLNOWRAP($cellContent['noWrap']);
                    }
                    if (isset($cellContent['cellMargin']) && is_array($cellContent['cellMargin'])) {
                        $this->generateCELLMARGIN($cellContent['cellMargin']);
                    } else if (isset($cellContent['cellMargin']) && !is_array($cellContent['cellMargin'])) {
                        $cellContent['cellMargins']['top'] = $cellContent['cellMargin'];
                        $cellContent['cellMargins']['left'] = $cellContent['cellMargin'];
                        $cellContent['cellMargins']['bottom'] = $cellContent['cellMargin'];
                        $cellContent['cellMargins']['right'] = $cellContent['cellMargin'];
                        $this->generateCELLMARGIN($cellContent['cellMargins']);
                    }
                    if (isset($cellContent['textDirection'])) {
                        $this->generateCELLTEXTDIRECTION($cellContent['textDirection']);
                    }
                    if (isset($cellContent['fitText'])) {
                        $this->generateCELLFITTEXT($cellContent['fitText']);
                    }
                    if (isset($cellContent['vAlign'])) {
                        $this->generateCELLVALIGN($cellContent['vAlign']);
                    }
                    $this->cleanTemplateTcPr();

                    /* if ($cellContent['value'] instanceof CreateText) {
                      if (!empty($cellContent['value']->_embeddedText[0]['cell_color'])) {
                      $this->generateSHD(
                      'solid',
                      $cellContent['value']->_embeddedText[0]['cell_color']
                      );
                      }
                      } else {
                      $this->generateSHD();
                      } */
                    if ($cellContent['value'] instanceof WordFragment) {
                        $this->_xml = str_replace('__PHX=__GENERATETC__', (string) $cellContent['value'], $this->_xml);
                    } else {
                        $this->generateP($cellContent['value'], $args[1]);
                    }
                    $colNumber++;
                }
                $this->cleanTemplateR();
            }
        }
    }

    /**
     * Add list
     *
     * @param string $list
     * @access protected
     */
    protected function addList($list)
    {
        $this->_xml = str_replace('__PHX=__GENERATEP__', $list, $this->_xml);
    }

    /**
     * Generate w:cantSplit
     *
     * @param string $val
     * @access protected
     */
    protected function generateTRCANTSPLIT($val = 'true')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':cantSplit w:val="' . $val . '"/>__PHX=__GENERATETRPR__';
        $this->_xml = str_replace('__PHX=__GENERATETRPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:tcBorders
     *
     * @param array $border
     * @access protected
     */
    protected function generateCELLBORDER($border)
    {
        $xml = '<w:tcBorders>';
        foreach (self::$_borders as $key => $value) {
            if (isset($border[$value])) {
                if (isset($border[$value]['type'])) {
                    $border_type = $border[$value]['type'];
                } else {
                    $border_type = 'single';
                }
                if (isset($border[$value]['color'])) {
                    $border_color = $border[$value]['color'];
                } else {
                    $border_color = '000000';
                }
                if (isset($border[$value]['sz'])) {
                    $border_sz = $border[$value]['sz'];
                } else {
                    $border_sz = 6;
                }
                if (isset($border[$value]['spacing'])) {
                    $border_spacing = $border[$value]['spacing'];
                } else {
                    $border_spacing = 0;
                }
                $xml .= '<w:' . $value . ' w:val="' . $border_type . '" w:color="' . $border_color . '" w:sz="' . $border_sz . '" w:space="' . $border_spacing . '"/>';
            }
        }
        $xml .= '</w:tcBorders>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:gridSpan
     *
     * @param string $val
     * @access protected
     */
    protected function generateCELLGRIDSPAN($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':gridSpan ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:tcMar
     *
     * @param array $cellMargin
     * @access protected
     */
    protected function generateCELLMARGIN($cellMargin)
    {
        $sides = array('top', 'left', 'bottom', 'right');
        $xml = '<w:tcMar>';
        foreach ($cellMargin as $key => $value) {
            if (in_array($key, $sides)) {
                $xml .= '<w:' . $key . ' w:w="' . $value . '" w:type="dxa" />';
            }
        }
        $xml .= '</w:tcMar>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:noWrap
     *
     * @param string $val
     * @access protected
     */
    protected function generateCELLNOWRAP($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':noWrap ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:shd
     *
     * @param string $color
     * @access protected
     */
    protected function generateCELLSHD($color)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':shd  w:val ="clear" w:fill="' . $color . '"/>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:tcFitText
     *
     * @param string $val
     * @access protected
     */
    protected function generateCELLFITTEXT($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tcFitText ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:textDirection
     *
     * @param string $val
     * @access protected
     */
    protected function generateCELLTEXTDIRECTION($val)
    {
        $textDirections = array('btLr', 'tbRl', 'lrTb', 'tbRl', 'btLr', 'lrTbV', 'tbRlV', 'tbLrV');
        if (in_array($val, $textDirections)) {
            $xml = '<' . CreateElement::NAMESPACEWORD .
                    ':textDirection ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATETCPR__';
            $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
        }
    }

    /**
     * Generate w:vAlign
     *
     * @param string $val
     * @access protected
     */
    protected function generateCELLVALIGN($val)
    {
        $valign = array('top', 'center', 'both', 'bottom');
        if (in_array($val, $valign)) {
            $xml = '<' . CreateElement::NAMESPACEWORD .
                    ':vAlign ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATETCPR__';
            $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
        }
    }

    /**
     * Generate w:vmerge
     *
     * @param string $val
     * @access protected
     */
    protected function generateCELLVMERGE($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':vMerge ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:vmerge
     *
     * @param mixed $val
     * @access protected
     */
    protected function generateCELLWIDTH($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tcW ' . CreateElement::NAMESPACEWORD . ':w="' . $val . '" w:type="dxa" />__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:gridcols
     *
     * @param string $w
     * @access protected
     */
    protected function generateGRIDCOLS($w)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblGrid ' .
                CreateElement::NAMESPACEWORD . ':w="' . $w . '"></' .
                CreateElement::NAMESPACEWORD . ':tblGrid>__PHX=__GENERATEGRIDCOLS__';
        $this->_xml = str_replace('__PHX=__GENERATEGRIDCOLS__', $xml, $this->_xml);
    }

    /**
     * Generate w:gridcol
     *
     * @access protected
     */
    protected function generateGRIDCOL($width = 1)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':gridCol';
        if (!empty($width)) {
            $xml .= ' w:w="' . $width . '" ';
        }
        $xml .= '/>__PHX=__GENERATEGRIDCOL__';
        $this->_xml = str_replace('__PHX=__GENERATEGRIDCOL__', $xml, $this->_xml);
    }

    /**
     * Generate w:hmerge
     *
     * @access protected
     * @deprecated
     * @param string $val
     */
    protected function generateHMERGE($val = '')
    {

    }

    /**
     * Generate w:jc
     *
     * @param string $val
     * @access protected
     */
    protected function generateJC($val = '')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':jc ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':jc>';
        $this->_xml = str_replace('__PHX=__GENERATEJC__', $xml, $this->_xml);
    }

    /**
     * Generate w:p
     *
     * @access protected
     */
    protected function generateP($value = null, $options = null)
    {
        if (!is_array($options)) {
            $options = array();
        }
        if (!isset($options['textProperties'])) {
            $options['textProperties'] = array();
        }

        $xmlWF = new WordFragment();
        $xmlWF->addText($value, $options['textProperties']);
        $xml = (string) $xmlWF;
        $this->_xml = str_replace('__PHX=__GENERATETC__', $xml, $this->_xml);
    }

    /**
     * Generate w:shd
     *
     * @param string $val
     * @param string $color
     * @param string $fill
     * @param string $bgcolor
     * @param string $themeFill
     * @access protected
     */
    protected function generateSHD($val = 'horz-cross', $color = 'ff0000', $fill = '', $bgcolor = '', $themeFill = '')
    {
        $xmlAux = '<' . CreateElement::NAMESPACEWORD . ':shd w:val="' .
                $val . '"';
        if ($color != '')
            $xmlAux .= ' w:color="' . $color . '"';
        if ($fill != '')
            $xmlAux .= ' w:fill="' . $fill . '"';
        if ($bgcolor != '')
            $xmlAux .= ' wx:bgcolor="' . $bgcolor . '"';
        if ($themeFill != '')
            $xmlAux .= ' w:themeFill="' . $themeFill . '"';
        $xmlAux .= '></' . CreateElement::NAMESPACEWORD . ':shd>';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xmlAux, $this->_xml);
    }

    /**
     * Generate w:tableLayout
     *
     * @param string $val
     * @access protected
     */
    protected function generateTBLLAYOUT($val = 'autofit')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblLayout ' .
                CreateElement::NAMESPACEWORD . ':type="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':tblLayout>';
        $this->_xml = str_replace('__PHX=__GENERATETBLLAYOUT__', $xml, $this->_xml);
    }

    /**
     * Generate w:bidiVisual
     *
     * @param mixed $val
     * @access protected
     */
    protected function generateTBLBIDI($val)
    {
        $xml = '<w:bidiVisual w:val="' . $val . '" />';
        $this->_xml = str_replace('__PHX=__GENERATETBLBIDI__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblborders
     *
     * @access protected
     */
    protected function generateTBLBORDERS()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblBorders>__PHX=__GENERATETBLBORDER__</' .
                CreateElement::NAMESPACEWORD . ':tblBorders>';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDERS__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblbottom
     *
     * @param string $val
     * @param string $sz
     * @param string $space
     * @param string $color
     * @access protected
     */
    protected function generateTBLBOTTOM($val = "single", $sz = "4", $space = '0', $color = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':bottom ' . CreateElement::NAMESPACEWORD .
                ':val="' . $val . '" ' . CreateElement::NAMESPACEWORD .
                ':sz="' . $sz . '" ' . CreateElement::NAMESPACEWORD .
                ':space="' . $space . '" ' . CreateElement::NAMESPACEWORD .
                ':color="' . $color . '"></' . CreateElement::NAMESPACEWORD .
                ':bottom>__PHX=__GENERATETBLBORDER__';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDER__', $xml, $this->_xml);
    }

    /**
     * Generate w:tbl
     *
     * @access protected
     */
    protected function generateTBL()
    {
        $this->_xml = '<' . CreateElement::NAMESPACEWORD .
                ':tbl>__PHX=__GENERATETBL__</' . CreateElement::NAMESPACEWORD .
                ':tbl>';
    }

    /**
     * Generate w:tblpPr
     *
     * @param array $float
     * @access protected
     */
    protected function generateTBLFLOAT($float)
    {
        $margin = array();
        foreach (self::$_borders as $value) {
            if (isset($float['textMargin_' . $value])) {
                $margin[$value] = (int) $float['textMargin_' . $value];
            } else {
                $margin[$value] = 0;
            }
        }
        if (isset($float['align'])) {
            $align = $float['align'];
        } else {
            $align = 'left';
        }

        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblpPr ';
        $xml .= 'w:leftFromText="' . $margin['left'] . '" ';
        $xml .= 'w:rightFromText="' . $margin['right'] . '" ';
        $xml .= 'w:topFromText="' . $margin['top'] . '" ';
        $xml .= 'w:bottomFromText="' . $margin['bottom'] . '" ';
        $xml .= 'w:vertAnchor="text" w:horzAnchor ="margin" ';
        $xml .= 'w:tblpXSpec ="' . $align . '" w:tblpYSpec="inside" />';

        $this->_xml = str_replace('__PHX=__GENERATETBLFLOAT__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblstyle
     *
     * @param string $strVal
     * @access protected
     */
    protected function generateTBLSTYLE($strVal = 'TableGridPHPDOCX')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblStyle ' . CreateElement::NAMESPACEWORD .
                ':val="' . $strVal . '"></' . CreateElement::NAMESPACEWORD .
                ':tblStyle>';
        $this->_xml = str_replace('__PHX=__GENERATETBLSTYLE__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblcellspacing
     *
     * @param string $w
     * @param string $type
     * @access protected
     */
    protected function generateTBLCELLSPACING($w = '', $type = 'dxa')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblCellSpacing ' . CreateElement::NAMESPACEWORD .
                ':w="' . $w . '" ' . CreateElement::NAMESPACEWORD .
                ':type="' . $type . '"></' . CreateElement::NAMESPACEWORD .
                ':tblCellSpacing>';
        $this->_xml = str_replace(
                '__PHX=__GENERATETBLCELLSPACING__', $xml, $this->_xml
        );
    }

    /**
     * Generate w:tblInd
     *
     * @param string $w
     * @param string $type
     * @access protected
     */
    protected function generateTBLINDENT($w = '', $type = 'dxa')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblInd ' . CreateElement::NAMESPACEWORD .
                ':w="' . $w . '" ' . CreateElement::NAMESPACEWORD .
                ':type="' . $type . '"></' . CreateElement::NAMESPACEWORD .
                ':tblInd>';
        $this->_xml = str_replace(
                '__PHX=__GENERATETBLINDENT__', $xml, $this->_xml
        );
    }

    /**
     * Generate w:tblCellMar
     *
     * @param array $cellMargin
     * @access protected
     */
    protected function generateTBLCELLMARGIN($cellMargin)
    {
        if (!is_array($cellMargin)) {
            $cellMargins['top'] = $cellMargin;
            $cellMargins['left'] = $cellMargin;
            $cellMargins['bottom'] = $cellMargin;
            $cellMargins['right'] = $cellMargin;
        } else {
            $cellMargins = $cellMargin;
        }
        $sides = array('top', 'left', 'bottom', 'right');
        $xml = '<w:tblCellMar>';
        foreach ($cellMargins as $key => $value) {
            if (in_array($key, $sides)) {
                $xml .= '<w:' . $key . ' w:w="' . $value . '" w:type="dxa" />';
            }
        }
        $xml .= '</w:tblCellMar>';
        $this->_xml = str_replace(
                '__PHX=__GENERATETBLCELLMARGINS__', $xml, $this->_xml
        );
    }

    /**
     * Generate w:tblgrid
     *
     * @access protected
     */
    protected function generateTBLGRID()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblGrid>__PHX=__GENERATEGRIDCOL__</' .
                CreateElement::NAMESPACEWORD .
                ':tblGrid>__PHX=__GENERATETBL__';
        $this->_xml = str_replace('__PHX=__GENERATETBL__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblinsideh
     *
     * @param string $val
     * @param string $sz
     * @param string $space
     * @param string $color
     * @access protected
     */
    protected function generateTBLINSIDEH($val = "single", $sz = "4", $space = '0', $color = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':insideH ' . CreateElement::NAMESPACEWORD .
                ':val="' . $val . '" ' . CreateElement::NAMESPACEWORD .
                ':sz="' . $sz . '" ' . CreateElement::NAMESPACEWORD .
                ':space="' . $space . '" ' . CreateElement::NAMESPACEWORD .
                ':color="' . $color . '"></' . CreateElement::NAMESPACEWORD .
                ':insideH>__PHX=__GENERATETBLBORDER__';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDER__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblinsidev
     *
     * @param string $val
     * @param string $sz
     * @param string $space
     * @param string $color
     * @access protected
     */
    protected function generateTBLINSIDEV($val = "single", $sz = "4", $space = '0', $color = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':insideV ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" ' .
                CreateElement::NAMESPACEWORD . ':sz="' . $sz . '" ' .
                CreateElement::NAMESPACEWORD . ':space="' . $space . '" ' .
                CreateElement::NAMESPACEWORD . ':color="' . $color . '"></' .
                CreateElement::NAMESPACEWORD . ':insideV>__PHX=__GENERATETBLBORDER__';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDER__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblleft
     *
     * @param string $val
     * @param string $sz
     * @param string $space
     * @param string $color
     * @access protected
     */
    protected function generateTBLLEFT($val = "single", $sz = "4", $space = '0', $color = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':left ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" ' .
                CreateElement::NAMESPACEWORD . ':sz="' . $sz . '" ' .
                CreateElement::NAMESPACEWORD . ':space="' . $space . '" ' .
                CreateElement::NAMESPACEWORD . ':color="' . $color . '"></' .
                CreateElement::NAMESPACEWORD . ':left>__PHX=__GENERATETBLBORDER__';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDER__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblLook
     *
     * @param array $conditionalFormatting
     * @access protected
     */
    protected function generateTBLLOOK($conditionalFormatting)
    {
        //0x0020=Apply first row conditional formatting
        //0x0040=Apply last row conditional formatting
        //0x0080=Apply first column conditional formatting
        //0x0100=Apply last column conditional formatting
        //0x0200=Do not apply row banding conditional formatting
        //0x0400=Do not apply column banding conditional formatting

        $mask = array();
        $mask['firstRow'] = 0x0020;
        $mask['lastRow'] = 0x0040;
        $mask['firstCol'] = 0x0080;
        $mask['lastCol'] = 0x0100;
        $mask['noHBand'] = 0x0200;
        $mask['noVBand'] = 0x0400;

        $result = 0;

        foreach ($conditionalFormatting as $key => $value) {
            if (!empty($value)) {
                $result |= $mask[$key];
            }
        }
        $valHex = dechex($result);
        $val = (string) $valHex;
        if (strlen($valHex) == 1) {
            $val = '000' . (string) $valHex;
        } else if (strlen($valHex) == 2) {
            $val = '00' . (string) $valHex;
        } else if (strlen($valHex) == 3) {
            $val = '0' . (string) $valHex;
        } else if (strlen($valHex) == 4) {
            $val = (string) $valHex;
        }

        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblLook ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" ></' .
                CreateElement::NAMESPACEWORD . ':tblLook>';
        $this->_xml = str_replace('__PHX=__GENERATETBLLOOK__', $xml, $this->_xml);
    }

    /**
     * Generate w:tbloverlap
     *
     * @param string $val
     * @access protected
     */
    protected function generateTBLOVERLAP($val = 'never')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblOverlap ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':tblOverlap>';
        $this->_xml = str_replace('__PHX=__GENERATETBLOVERLAP__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblCaption
     *
     * @param string $val
     * @access protected
     */
    protected function generateTBLDESCR($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblCaption ' .
                CreateElement::NAMESPACEWORD . ':val="' . $this->parseAndCleanTextString($val) . '"></' .
                CreateElement::NAMESPACEWORD . ':tblCaption>';
        $this->_xml = str_replace('__PHX=__GENERATETBLDESCRTITLE__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblDescription
     *
     * @param string $val
     * @access protected
     */
    protected function generateTBLDESCRTITLE($val)
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblDescription ' .
                CreateElement::NAMESPACEWORD . ':val="' . $this->parseAndCleanTextString($val) . '"></' .
                CreateElement::NAMESPACEWORD . ':tblDescription>';
        $this->_xml = str_replace('__PHX=__GENERATETBLDESCRIPTION__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblpr
     *
     * @access protected
     */
    protected function generateTBLPR()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblPr>__PHX=__GENERATETBLSTYLE____PHX=__GENERATETBLFLOAT____PHX=__GENERATETBLOVERLAP____PHX=__GENERATETBLBIDI____PHX=__GENERATETBLW__' .
                '__PHX=__GENERATEJC____PHX=__GENERATETBLCELLSPACING____PHX=__GENERATETBLINDENT__' .
                '__PHX=__GENERATETBLBORDERS____PHX=__GENERATESHD____PHX=__GENERATETBLLAYOUT____PHX=__GENERATETBLCELLMARGINS____PHX=__GENERATETBLLOOK____PHX=__GENERATETBLDESCRTITLE____PHX=__GENERATETBLDESCRIPTION__</' .
                CreateElement::NAMESPACEWORD . ':tblPr>__PHX=__GENERATETBL__';
        $this->_xml = str_replace('__PHX=__GENERATETBL__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblright
     *
     * @param string $val
     * @param string $sz
     * @param string $space
     * @param string $color
     * @access protected
     */
    protected function generateTBLRIGHT($val = 'single', $sz = '4', $space = '0', $color = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':right ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" ' .
                CreateElement::NAMESPACEWORD . ':sz="' . $sz . '" ' .
                CreateElement::NAMESPACEWORD . ':space="' . $space . '" ' .
                CreateElement::NAMESPACEWORD . ':color="' . $color . '"></' .
                CreateElement::NAMESPACEWORD . ':right>__PHX=__GENERATETBLBORDER__';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDER__', $xml, $this->_xml);
    }

    /**
     * Generate w:tbltop
     *
     * @param string $val
     * @param string $sz
     * @param string $space
     * @param string $color
     * @access protected
     */
    protected function generateTBLTOP($val = 'single', $sz = '4', $space = '0', $color = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':top ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" ' .
                CreateElement::NAMESPACEWORD . ':sz="' . $sz . '" ' .
                CreateElement::NAMESPACEWORD . ':space="' . $space . '" ' .
                CreateElement::NAMESPACEWORD . ':color="' . $color . '"></' .
                CreateElement::NAMESPACEWORD . ':top>__PHX=__GENERATETBLBORDER__';
        $this->_xml = str_replace('__PHX=__GENERATETBLBORDER__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblw
     *
     * @param string $type
     * @param mixed $value
     * @access protected
     */
    protected function generateTBLW($type, $value)
    {
        if ($type == 'pct') {
            $value = $value * 50;
        }
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tblW ' .
                CreateElement::NAMESPACEWORD . ':w="' . $value . '" ' .
                CreateElement::NAMESPACEWORD . ':type="' . $type . '"></' .
                CreateElement::NAMESPACEWORD . ':tblW>';
        $this->_xml = str_replace('__PHX=__GENERATETBLW__', $xml, $this->_xml);
    }

    /**
     * Generate w:tc
     *
     * @access protected
     */
    protected function generateTC()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tc >__PHX=__GENERATETC__</' .
                CreateElement::NAMESPACEWORD . ':tc>__PHX=__GENERATETR__';
        $this->_xml = str_replace('__PHX=__GENERATETR__', $xml, $this->_xml);
    }

    /**
     * Generate w:tcpr
     *
     * @access protected
     */
    protected function generateTCPR()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tcPr>__PHX=__GENERATETCPR__</' . CreateElement::NAMESPACEWORD .
                ':tcPr>__PHX=__GENERATETC__';
        $this->_xml = str_replace('__PHX=__GENERATETC__', $xml, $this->_xml);
    }

    /**
     * Generate w:tcw
     *
     * @param string $w
     * @param string $type
     * @access protected
     */
    protected function generateTCW($w = '', $type = 'dxa')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':tcW ' .
                CreateElement::NAMESPACEWORD . ':w="' . $w . '" ' .
                CreateElement::NAMESPACEWORD . ':type="' . $type . '"></' .
                CreateElement::NAMESPACEWORD . ':tcW>__PHX=__GENERATETCPR__';
        $this->_xml = str_replace('__PHX=__GENERATETCPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:tr
     *
     * @access protected
     */
    protected function generateTR()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tr >__PHX=__GENERATETRPR____PHX=__GENERATETR__</' .
                CreateElement::NAMESPACEWORD . ':tr>__PHX=__GENERATETBL__';
        $this->_xml = str_replace('__PHX=__GENERATETBL__', $xml, $this->_xml);
    }

    /**
     * Generate w:trheight
     *
     * @param mixed $height
     * @param string $hRule
     * @access protected
     */
    protected function generateTRHEIGHT($height = '', $hRule = 'auto')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':trHeight ' .
                CreateElement::NAMESPACEWORD . ':val="' . (int) $height .
                '" w:hRule="' . $hRule . '"></' .
                CreateElement::NAMESPACEWORD . ':trHeight>__PHX=__GENERATETRPR__';
        $this->_xml = str_replace('__PHX=__GENERATETRPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:trpr
     *
     * @access protected
     */
    protected function generateTRPR()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':trPr>__PHX=__GENERATETRPR__</' . CreateElement::NAMESPACEWORD .
                ':trPr>';
        $this->_xml = str_replace('__PHX=__GENERATETRPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:tblHeader
     *
     * @access protected
     */
    protected function generateTRTABLEHEADER()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD .
                ':tblHeader />__PHX=__GENERATETRPR__';
        $this->_xml = str_replace('__PHX=__GENERATETRPR__', $xml, $this->_xml);
    }

    /**
     * Clean template r token
     *
     * @access private
     */
    private function cleanTemplateR()
    {
        $this->_xml = preg_replace('/__PHX=__GENERATETR__/', '', $this->_xml);
    }

    /**
     * Clean template trpr token
     *
     * @access private
     */
    private function cleanTemplateTrPr()
    {
        $this->_xml = preg_replace('/__PHX=__GENERATETRPR__/', '', $this->_xml);
    }

    /**
     * Clean template trpr token
     *
     * @access private
     */
    private function cleanTemplateTcPr()
    {
        $this->_xml = preg_replace('/__PHX=__GENERATETCPR__/', '', $this->_xml);
    }

    /**
     * Prepares the table data for insertion
     * @param array $tableData
     * @access private
     */
    private function parseTableData($tableData)
    {
        $parsedData = array();
        $colCount = array();
        foreach ($tableData as $rowNumber => $row) {
            $parsedData[$rowNumber] = array();
            $colNumber = 0;
            foreach ($row as $col => $cell) {
                //Check if in the previous row there was a cell with rowspan > 1
                while (isset($parsedData[$rowNumber - 1][$colNumber]['rowspan']) &&
                $parsedData[$rowNumber - 1][$colNumber]['rowspan'] > 1) {
                    //replicate the array
                    $parsedData[$rowNumber][$colNumber] = $parsedData[$rowNumber - 1][$colNumber];
                    //reduce by one the rowspan
                    $parsedData[$rowNumber][$colNumber]['rowspan'] = $parsedData[$rowNumber - 1][$colNumber]['rowspan'] - 1;
                    //set up the vmerge and content values
                    $parsedData[$rowNumber][$colNumber]['vmerge'] = 'continue';
                    $parsedData[$rowNumber][$colNumber]['value'] = NULL;
                    if (isset($parsedData[$rowNumber - 1][$colNumber]['colspan'])) {
                        $colNumber += $parsedData[$rowNumber - 1][$colNumber]['colspan'];
                    } else {
                        $colNumber++;
                    }
                }
                if (is_array($cell)) {
                    $parsedData[$rowNumber][$colNumber] = $cell;
                } else {
                    $parsedData[$rowNumber][$colNumber]['value'] = $cell;
                }
                if (isset($parsedData[$rowNumber][$colNumber]['colspan'])) {
                    $colNumber += $parsedData[$rowNumber][$colNumber]['colspan'];
                } else {
                    $colNumber++;
                }
            }
            //check that there are no trailing rawspans after running through all cols
            if ($rowNumber > 0) {
                $colDiff = $colCount[$rowNumber - 1] - $colNumber;
                if ($colDiff > 0) {
                    for ($k = 0; $k < $colDiff; $k++) {
                        //Check if in the previous row there was a cell with rowspan > 1
                        while (isset($parsedData[$rowNumber - 1][$colNumber]['rowspan']) &&
                        $parsedData[$rowNumber - 1][$colNumber]['rowspan'] > 1) {
                            //replicate the array
                            $parsedData[$rowNumber][$colNumber] = $parsedData[$rowNumber - 1][$colNumber];
                            //reduce by one the rowspan
                            $parsedData[$rowNumber][$colNumber]['rowspan'] = $parsedData[$rowNumber - 1][$colNumber]['rowspan'] - 1;
                            //set up the vmerge and content values
                            $parsedData[$rowNumber][$colNumber]['vmerge'] = 'continue';
                            $parsedData[$rowNumber][$colNumber]['value'] = NULL;
                            if (isset($parsedData[$rowNumber - 1][$colNumber]['colspan'])) {
                                $colNumber += $parsedData[$rowNumber - 1][$colNumber]['colspan'];
                            } else {
                                $colNumber++;
                            }
                        }
                    }
                }
            }
            $colCount[$rowNumber] = $colNumber;
        }
        return $parsedData;
    }
}