<?php

/**
 * Create text
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateText extends CreateElement
{
    const IDTITLE = 229998237;

    /**
     *
     * @var string
     */
    protected $_embeddedText;

    /**
     *
     * @access private
     * @var array
     * @static
     */
    private static $_borders = array('top', 'left', 'bottom', 'right');

    /**
     *
     * @access private
     * @var mixed
     */
    private static $_instance = NULL;

    /**
     *
     * @access private
     * @var int
     */
    private static $_idTitle = 0;

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
     * Magic method, returns current XML
     *
     * @access public
     * @return string Return current XML
     */
    public function __toString()
    {
        return $this->_xml;
    }

    /**
     * Singleton, return instance of class
     *
     * @access public
     * @return CreateText
     * @static
     */
    public static function getInstance()
    {
        if (self::$_instance == NULL) {
            self::$_instance = new CreateText();
        }
        return self::$_instance;
    }

    /**
     * Create text
     *
     * @access public
     */
    public function createText()
    {
        $this->_xml = '';
        $args = func_get_args();

        $args[1] = CreateDocx::translateTextOptions2StandardFormat($args[1]);

        $this->generateP();

        $this->generatePPR();
        if (!empty($args[1]['pStyle'])) {
            $this->generatePSTYLE($args[1]['pStyle']);
        }
        if (!empty($args[1]['keepNext'])) {
            $this->generateKEEPNEXT($args[1]['keepNext']);
        }
        if (!empty($args[1]['keepLines'])) {
            $this->generateKEEPLINES($args[1]['keepLines']);
        }
        if (!empty($args[1]['suppressLineNumbers'])) {
            $this->generateSUPPRESSLINENUMBERS($args[1]['suppressLineNumbers']);
        }
        if (!empty($args[1]['pageBreakBefore'])) {
            $this->generatePAGEBREAKBEFORE($args[1]['pageBreakBefore']);
        }
        if (!empty($args[1]['widowControl'])) {
            $this->generateWIDOWCONTROL($args[1]['widowControl']);
        }
        //we set the drawPBorders to false
        $drawPBorders = false;
        $border = array();

        //Run over the general border properties
        if (isset($args[1]['border'])) {
            $drawPBorders = true;
            foreach (self::$_borders as $key => $value) {
                $border[$value]['type'] = $args[1]['border'];
            }
        }
        if (isset($args[1]['border_color'])) {
            $drawPBorders = true;
            foreach (self::$_borders as $key => $value) {
                $border[$value]['color'] = $args[1]['border_color'];
            }
        }
        if (isset($args[1]['border_spacing'])) {
            $drawPBorders = true;
            foreach (self::$_borders as $key => $value) {
                $border[$value]['spacing'] = $args[1]['border_spacing'];
            }
        }
        if (isset($args[1]['border_sz'])) {
            $drawPBorders = true;
            foreach (self::$_borders as $key => $value) {
                $border[$value]['sz'] = $args[1]['border_sz'];
            }
        }
        //Run over the border choices of each side
        foreach (self::$_borders as $key => $value) {
            if (isset($args[1]['border_' . $value])) {
                $drawPBorders = true;
                $border[$value]['type'] = $args[1]['border_' . $value];
            }
            if (isset($args[1]['border_' . $value . '_color'])) {
                $drawPBorders = true;
                $border[$value]['color'] = $args[1]['border_' . $value . '_color'];
            }
            if (isset($args[1]['border_' . $value . '_spacing'])) {
                $drawPBorders = true;
                $border[$value]['spacing'] = $args[1]['border_' . $value . '_spacing'];
            }
            if (isset($args[1]['border_' . $value . '_sz'])) {
                $drawPBorders = true;
                $border[$value]['sz'] = $args[1]['border_' . $value . '_sz'];
            }
        }
        if ($drawPBorders) {
            $this->generatePBDR($border);
        }
        if (!empty($args[1]['backgroundColor'])) {
            $this->generatePPRSHD($args[1]['backgroundColor']);
        }
        if (!empty($args[1]['jc'])) {
            $this->generateJC($args[1]['jc']);
        }
        if (!empty($args[1]['tabPositions']) && is_array($args[1]['tabPositions'])) {
            $this->generateTABPOSITIONS($args[1]['tabPositions']);
        }
        if (!empty($args[1]['wordWrap'])) {
            $this->generateWORDWRAP($args[1]['wordWrap']);
        }
        if (!empty($args[1]['bidi'])) {
            $this->generateBIDI($args[1]['bidi']);
        }
        if (isset($args[1]['lineSpacing']) ||
                isset($args[1]['spacingTop']) ||
                isset($args[1]['spacingBottom'])) {
            if (!isset($args[1]['lineSpacing'])) {
                $args[1]['lineSpacing'] = 240;
            }
            if (!isset($args[1]['spacingTop'])) {
                $args[1]['spacingTop'] = '';
            }
            if (!isset($args[1]['spacingBottom'])) {
                $args[1]['spacingBottom'] = '';
            }
            $this->generateSPACING($args[1]['spacingTop'], $args[1]['spacingBottom'], $args[1]['lineSpacing']);
        }
        if (
                isset($args[1]['indent_left']) ||
                isset($args[1]['indent_right']) ||
                isset($args[1]['firstLineIndent']) ||
                isset($args[1]['hanging']
                )
        ) {
            $indentation = array();
            if (isset($args[1]['indent_left'])) {
                $indentation['left'] = $args[1]['indent_left'];
            }
            if (isset($args[1]['indent_right'])) {
                $indentation['right'] = $args[1]['indent_right'];
            }
            if (isset($args[1]['hanging'])) {
                $indentation['hanging'] = $args[1]['hanging'];
            }
            if (isset($args[1]['firstLineIndent'])) {
                $indentation['firstLine'] = $args[1]['firstLineIndent'];
            }


            $this->generateINDENT($indentation);
        }
        if (!empty($args[1]['contextualSpacing'])) {
            $this->generateCONTEXTUALSPACING($args[1]['contextualSpacing']);
        }
        if (!empty($args[1]['textDirection'])) {
            $this->generateTEXTDIRECTION($args[1]['textDirection']);
        }
        if (!empty($args[1]['headingLevel'])) {
            $this->generateHEADINGLEVEL($args[1]['headingLevel']);
        }
        //We include now paragraph wide run properties
        $this->generatePPRRPR();
        if (!empty($args[1]['rStyle'])) {
            $this->generateRSTYLE($args[1]['rStyle']);
        }
        if (!empty($args[1]['font'])) {
            $this->generateRFONTS($args[1]['font']);
        }
        if (!empty($args[1]['b'])) {
            $this->generateB($args[1]['b']);
        }
        if (!empty($args[1]['i'])) {
            $this->generateI($args[1]['i']);
        }
        if (!empty($args[1]['caps'])) {
            $this->generateCAPS($args[1]['caps']);
        }
        if (!empty($args[1]['position'])) {
            $this->generatePOSITION($args[1]['position']);
        }
        if (!empty($args[1]['scaling'])) {
            $this->generateSCALING($args[1]['scaling']);
        }
        if (!empty($args[1]['spacing'])) {
            $this->generateCHARACTERSPACING($args[1]['spacing']);
        }
        if (!empty($args[1]['superscript'])) {
            $this->generateVERTALIGN('superscript');
        }
        if (!empty($args[1]['subscript'])) {
            $this->generateVERTALIGN('subscript');
        }
        if (!empty($args[1]['strikeThrough'])) {
            $this->generateSTRIKETHROUGH();
        }
        if (!empty($args[1]['doubleStrikeThrough'])) {
            $this->generateDOUBLESTRIKETHROUGH();
        }
        if (!empty($args[1]['em'])) {
            $this->generateEM($args[1]['em']);
        }
        if (!empty($args[1]['smallCaps'])) {
            $this->generateSMALLCAPS($args[1]['smallCaps']);
        }
        if (!empty($args[1]['color'])) {
            $this->generateCOLOR($args[1]['color']);
        }
        if (!empty($args[1]['sz'])) {
            $this->generateSZ($args[1]['sz']);
        }
        if (!empty($args[1]['u'])) {
            if (!isset($args[1]['underlineColor'])) {
                $args[1]['underlineColor'] = null;
            }
            $this->generateU($args[1]['u'], $args[1]['underlineColor']);
        }
        if (!empty($args[1]['rtl'])) {
            $this->generateRTL($args[1]['rtl']);
        }
        if (!empty($args[1]['highlightColor'])) {
            $this->generateHIGHLIGHT($args[1]['highlightColor']);
        }
        if (!empty($args[1]['characterBorder'])) {
            $this->generateCHARACTERBORDER($args[1]['characterBorder']);
        }
        if (!empty($args[1]['tab']) && $args[1]['tab']) {
            $this->generateTABS();
        }
        if (!empty($args[1]['vanish'])) {
            $this->generateVANISH();
        }
        if (!empty($args[1]['emboss'])) {
            $this->generateEMBOSS();
        }
        if (!empty($args[1]['noProof'])) {
            $this->generateNOPROOF();
        }
        if (!empty($args[1]['outline'])) {
            $this->generateOUTLINE();
        }
        if (!empty($args[1]['shadow'])) {
            $this->generateSHADOW();
        }
        $this->cleanTemplateFirstRPR();

        if (!is_array($args[0])) {
            if ($args[0] instanceof WordFragment) {
                $runContent = $args[0]->inlineWordML();
                if (preg_match("/__PHX=__GENERATEP__/", $this->_xml)) {
                    $this->_xml = str_replace('__PHX=__GENERATEP__', $runContent . '__PHX=__GENERATESUBR__', $this->_xml);
                } else {
                    $this->_xml = str_replace('__PHX=__GENERATESUBR__', $runContent . '__PHX=__GENERATESUBR__', $this->_xml);
                }
            } else {
                $this->generateR();
                $this->generateRPR();
                if (!empty($args[1]['rStyle'])) {
                    $this->generateRSTYLE($args[1]['rStyle']);
                }
                if (!empty($args[1]['font'])) {
                    $this->generateRFONTS($args[1]['font']);
                }
                if (!empty($args[1]['b'])) {
                    $this->generateB($args[1]['b']);
                }
                if (!empty($args[1]['i'])) {
                    $this->generateI($args[1]['i']);
                }
                if (!empty($args[1]['caps'])) {
                    $this->generateCAPS($args[1]['caps']);
                }
                if (!empty($args[1]['position'])) {
                        $this->generatePOSITION($args[1]['position']);
                }
                if (!empty($args[1]['scaling'])) {
                    $this->generateSCALING($args[1]['scaling']);
                }
                if (!empty($args[1]['spacing'])) {
                    $this->generateCHARACTERSPACING($args[1]['spacing']);
                }
                if (!empty($args[1]['superscript'])) {
                    $this->generateVERTALIGN('superscript');
                }
                if (!empty($args[1]['subscript'])) {
                    $this->generateVERTALIGN('subscript');
                }
                if (!empty($args[1]['strikeThrough'])) {
                    $this->generateSTRIKETHROUGH();
                }
                if (!empty($args[1]['doubleStrikeThrough'])) {
                    $this->generateDOUBLESTRIKETHROUGH();
                }
                if (!empty($args[1]['em'])) {
                    $this->generateEM($args[1]['em']);
                }
                if (!empty($args[1]['lang']) && is_array($args[1]['lang'])) {
                    $this->generateLANG($args[1]['lang']);
                }
                if (!empty($args[1]['smallCaps'])) {
                    $this->generateSMALLCAPS($args[1]['smallCaps']);
                }
                if (!empty($args[1]['characterBorder'])) {
                    $this->generateCHARACTERBORDER($args[1]['characterBorder']);
                }
                if (!empty($args[1]['color'])) {
                    $this->generateCOLOR($args[1]['color']);
                }
                if (!empty($args[1]['sz'])) {
                    $this->generateSZ($args[1]['sz']);
                }
                if (!empty($args[1]['u'])) {
                    if (!isset($args[1]['underlineColor'])) {
                        $args[1]['underlineColor'] = null;
                    }
                    $this->generateU($args[1]['u'], $args[1]['underlineColor']);
                }
                if (!empty($args[1]['rtl'])) {
                    $this->generateRTL($args[1]['rtl']);
                }
                if (!empty($args[1]['highlightColor'])) {
                    $this->generateHIGHLIGHT($args[1]['highlightColor']);
                }
                if (!empty($args[1]['tab']) && $args[1]['tab']) {
                    $this->generateTABS();
                }
                if (!empty($args[1]['vanish']) && $args[1]['vanish']) {
                    $this->generateVANISH();
                }
                if (!empty($args[1]['emboss']) && $args[1]['emboss']) {
                    $this->generateEMBOSS();
                }
                if (!empty($args[1]['noProof']) && $args[1]['noProof']) {
                    $this->generateNOPROOF();
                }
                if (!empty($args[1]['outline']) && $args[1]['outline']) {
                    $this->generateOUTLINE();
                }
                if (!empty($args[1]['shadow']) && $args[1]['shadow']) {
                    $this->generateSHADOW();
                }
                if (empty($args[1]['spaces'])) {
                    $args[1]['spaces'] = '';
                }
                if (!isset($args[1]['lineBreak'])) {
                    $args[1]['lineBreak'] = false;
                }
                if (!isset($args[1]['columnBreak'])) {
                    $args[1]['columnBreak'] = false;
                }

                $this->generateT($args[0], $args[1]['spaces'], $args[1]['lineBreak'], $args[1]['columnBreak']);
                $this->cleanTemplateFirstRPR();
            }
        } else {
            foreach ($args[0] as $texts) {
                $texts = CreateDocx::translateTextOptions2StandardFormat($texts);
                if ($texts instanceof WordFragment) {
                    $runContent = $texts->inlineWordML();
                    if (preg_match("/__PHX=__GENERATEP__/", $this->_xml)) {
                        $this->_xml = str_replace('__PHX=__GENERATEP__', $runContent . '__PHX=__GENERATESUBR__', $this->_xml);
                    } else if (preg_match("/__PHX=__GENERATER__/", $this->_xml)) {
                        $this->_xml = str_replace('__PHX=__GENERATER__', $runContent . '__PHX=__GENERATER__', $this->_xml);
                    } else {
                        $this->_xml = str_replace('__PHX=__GENERATESUBR__', $runContent . '__PHX=__GENERATESUBR__', $this->_xml);
                    }
                } else {
                    $texts = CreateDocx::setRTLOptions($texts);
                    $this->generateR();

                    //Inherit run styles from paragraph styles if they have not been
                    //explicitely set
                    if (empty($texts['b']) && !empty($args[1]['b'])) {
                        $texts['b'] = $args[1]['b'];
                    }
                    if (empty($texts['i']) && !empty($args[1]['i'])) {
                        $texts['i'] = $args[1]['i'];
                    }
                    if (empty($texts['caps']) && !empty($args[1]['caps'])) {
                        $texts['caps'] = $args[1]['caps'];
                    }
                    if (empty($texts['position']) && !empty($args[1]['position'])) {
                        $texts['position'] = $args[1]['position'];
                    }
                    if (empty($texts['scaling']) && !empty($args[1]['scaling'])) {
                        $texts['scaling'] = $args[1]['scaling'];
                    }
                    if (empty($texts['spacing']) && !empty($args[1]['spacing'])) {
                        $texts['spacing'] = $args[1]['spacing'];
                    }
                    if (empty($texts['superscript']) && !empty($args[1]['superscript'])) {
                        $texts['superscript'] = $args[1]['superscript'];
                    }
                    if (empty($texts['subscript']) && !empty($args[1]['subscript'])) {
                        $texts['subscript'] = $args[1]['subscript'];
                    }
                    if (!empty($texts['strikeThrough']) && !empty($args[1]['strikeThrough'])) {
                        $texts['strikeThrough'] = $args[1]['strikeThrough'];
                    }
                    if (!empty($texts['doubleStrikeThrough']) && !empty($args[1]['doubleStrikeThrough'])) {
                        $texts['doubleStrikeThrough'] = $args[1]['doubleStrikeThrough'];
                    }
                    if (empty($texts['em']) && !empty($args[1]['em'])) {
                        $texts['em'] = $args[1]['em'];
                    }
                    if (empty($texts['smallCaps']) && !empty($args[1]['smallCaps'])) {
                        $texts['smallCaps'] = $args[1]['smallCaps'];
                    }
                    if (empty($texts['u']) && !empty($args[1]['u'])) {
                        $texts['u'] = $args[1]['u'];
                    }
                    if (empty($texts['underlineColor']) && !empty($args[1]['underlineColor'])) {
                        $texts['underlineColor'] = $args[1]['underlineColor'];
                    }
                    if (empty($texts['rtl']) && !empty($args[1]['rtl'])) {
                        $texts['rtl'] = $args[1]['rtl'];
                    }
                    if (empty($texts['highlightColor']) && !empty($args[1]['highlightColor'])) {
                        $texts['highlightColor'] = $args[1]['highlightColor'];
                    }
                    if (empty($texts['sz']) && !empty($args[1]['sz'])) {
                        $texts['sz'] = $args[1]['sz'];
                    }
                    if (empty($texts['characterBorder']) && !empty($args[1]['characterBorder'])) {
                        $texts['characterBorder'] = $args[1]['characterBorder'];
                    }
                    if (empty($texts['color']) && !empty($args[1]['color'])) {
                        $texts['color'] = $args[1]['color'];
                    }
                    if (empty($texts['rStyle']) && !empty($args[1]['rStyle'])) {
                        $texts['rStyle'] = $args[1]['rStyle'];
                    }
                    if (empty($texts['font']) && !empty($args[1]['font'])) {
                        $texts['font'] = $args[1]['font'];
                    }
                    if (empty($texts['tab']) && !empty($args[1]['tab'])) {
                        $texts['tab'] = $args[1]['tab'];
                    }
                    if (empty($texts['vanish']) && !empty($args[1]['vanish'])) {
                        $texts['vanish'] = $args[1]['vanish'];
                    }
                    if (empty($texts['emboss']) && !empty($args[1]['emboss'])) {
                        $texts['emboss'] = $args[1]['emboss'];
                    }
                    if (empty($texts['noProof']) && !empty($args[1]['noProof'])) {
                        $texts['noProof'] = $args[1]['noProof'];
                    }
                    if (empty($texts['outline']) && !empty($args[1]['outline'])) {
                        $texts['outline'] = $args[1]['outline'];
                    }
                    if (empty($texts['shadow']) && !empty($args[1]['shadow'])) {
                        $texts['shadow'] = $args[1]['shadow'];
                    }

                    $this->generateRPR();
                    if (!empty($texts['rStyle'])) {
                        $this->generateRSTYLE($texts['rStyle']);
                    }
                    if (!empty($texts['font'])) {
                        $this->generateRFONTS($texts['font']);
                    }
                    if (!empty($texts['b'])) {
                        $this->generateB($texts['b']);
                    }
                    if (!empty($texts['i'])) {
                        $this->generateI($texts['i']);
                    }
                    if (!empty($texts['caps'])) {
                        $this->generateCAPS($texts['caps']);
                    }
                    if (!empty($texts['position'])) {
                        $this->generatePOSITION($texts['position']);
                    }
                    if (!empty($texts['scaling'])) {
                        $this->generateSCALING($texts['scaling']);
                    }
                    if (!empty($texts['spacing'])) {
                        $this->generateCHARACTERSPACING($texts['spacing']);
                    }
                    if (!empty($texts['superscript'])) {
                        $this->generateVERTALIGN('superscript');
                    }
                    if (!empty($texts['subscript'])) {
                        $this->generateVERTALIGN('subscript');
                    }
                    if (!empty($texts['strikeThrough'])) {
                        $this->generateSTRIKETHROUGH();
                    }
                    if (!empty($texts['doubleStrikeThrough'])) {
                        $this->generateDOUBLESTRIKETHROUGH();
                    }
                    if (!empty($texts['em'])) {
                        $this->generateEM($texts['em']);
                    }
                    if (!empty($texts['lang']) && is_array($texts['lang'])) {
                        $this->generateLANG($texts['lang']);
                    }
                    if (!empty($texts['smallCaps'])) {
                        $this->generateSMALLCAPS($texts['smallCaps']);
                    }
                    if (!empty($texts['u'])) {
                        if (!isset($texts['underlineColor'])) {
                            $texts['underlineColor'] = null;
                        }
                        $this->generateU($texts['u'], $texts['underlineColor']);
                    }
                    if (!empty($texts['rtl'])) {
                        $this->generateRTL($texts['rtl']);
                    }
                    if (!empty($texts['highlightColor'])) {
                        $this->generateHIGHLIGHT($texts['highlightColor']);
                    }
                    if (!empty($texts['characterBorder'])) {
                        $this->generateCHARACTERBORDER($texts['characterBorder']);
                    }
                    if (!empty($texts['sz'])) {
                        $this->generateSZ($texts['sz']);
                    }
                    if (!empty($texts['color'])) {
                        $this->generateCOLOR($texts['color']);
                    }
                    if (!empty($texts['tab']) && $texts['tab']) {
                        $this->generateTABS();
                    }
                    if (empty($texts['spaces'])) {
                        $texts['spaces'] = '';
                    }
                    if (!isset($texts['lineBreak'])) {
                        $texts['lineBreak'] = false;
                    }
                    if (!isset($texts['columnBreak'])) {
                        $texts['columnBreak'] = false;
                    }
                    if (!isset($texts['text'])) {
                        $texts['text'] = '';
                    }
                    if (!empty($texts['vanish'])) {
                        $this->generateVANISH();
                    }
                    if (!empty($texts['emboss'])) {
                        $this->generateEMBOSS();
                    }
                    if (!empty($texts['noProof'])) {
                        $this->generateNOPROOF();
                    }
                    if (!empty($texts['outline'])) {
                        $this->generateOUTLINE();
                    }
                    if (!empty($texts['shadow'])) {
                        $this->generateSHADOW();
                    }

                    $this->generateT($texts['text'], $texts['spaces'], $texts['lineBreak'], $texts['columnBreak']);
                    $this->cleanTemplateFirstRPR();
                }
            }
        }
    }

    /**
     * Create title
     *
     * @access protected
     */
    public function createTitle()
    {
        $this->_xml = '';
        $args = func_get_args();
        if (isset($args[0])) {
            $this->generateP();
            $this->generatePPR();
            if (!isset($args[1]['val'])) {
                $args[1]['val'] = '';
            }
            if (isset($args[1]['type'])) {
                if ($args[1]['type'] == 'subtitle') {
                    $this->generatePSTYLE('SubtitlePHPDOCX');
                } else {
                    $this->generatePSTYLE('TitlePHPDOCX');
                }
            } else {
                $this->generatePSTYLE('TitlePHPDOCX');
            }
            if (isset($args[1]['pageBreakBefore'])) {
                $this->generatePAGEBREAKBEFORE($args[1]['pageBreakBefore']);
            }
            if (isset($args[1]['widowControl'])) {
                $this->generateWIDOWCONTROL($args[1]['widowControl']);
            }
            if (!empty($args[1]['tabPositions']) && is_array($args[1]['tabPositions'])) {
                $this->generateTABPOSITIONS($args[1]['tabPositions']);
            }
            if (isset($args[1]['wordWrap'])) {
                $this->generateWORDWRAP($args[1]['wordWrap']);
            }
            if (isset($args[1]['bidi'])) {
                $this->generateBIDI($args[1]['bidi']);
            }
            self::$_idTitle++;
            $this->generateBOOKMARKSTART(
                    self::$_idTitle, '_Toc' . (self::$_idTitle + self::IDTITLE)
            );
            $this->generateR();
            if (
                    isset($args[1]['b']) ||
                    isset($args[1]['i']) ||
                    isset($args[1]['u']) ||
                    isset($args[1]['rtl']) ||
                    isset($args[1]['sz']) ||
                    isset($args[1]['color']) ||
                    isset($args[1]['font'])
            ) {
                $this->generateRPR();
                if (isset($args[1]['font'])) {
                    $this->generateRFONTS($args[1]['font']);
                }
                if (isset($args[1]['b'])) {
                    $this->generateB($args[1]['b']);
                }
                if (isset($args[1]['i'])) {
                    $this->generateI($args[1]['i']);
                }
                if (isset($args[1]['color'])) {
                    $this->generateCOLOR($args[1]['color']);
                }
                if (isset($args[1]['sz'])) {
                    $this->generateSZ($args[1]['sz']);
                }
                if (isset($args[1]['u'])) {
                    if (!isset($args[1]['underlineColor'])) {
                        $args[1]['underlineColor'] = null;
                    }
                    $this->generateU($args[1]['u'], $args[1]['underlineColor']);
                }
                if (isset($args[1]['rtl'])) {
                    $this->generateRTL($args[1]['rtl']);
                }
            }
            $this->generateT($args[0]);
            $this->generateBOOKMARKEND(self::$_idTitle);
            $this->cleanTemplate();
        }
    }

    /**
     * Init text
     *
     * @access public
     */
    public function initText()
    {
        $args = func_get_args();

        $this->_embeddedText = $args[0];
    }

    /**
     * Generate w:bidi
     *
     * @access protected
     * @param string $val
     */
    protected function generateBIDI($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':bidi w:val="' . $val . '"></' . CreateElement::NAMESPACEWORD .
                ':bidi>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:bookmarkend
     *
     * @access protected
     * @param int $id
     */
    protected function generateBOOKMARKEND($id)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEBOOKMARKEND__', '<' . CreateElement::NAMESPACEWORD .
                ':bookmarkEnd ' . CreateElement::NAMESPACEWORD . ':id="' . $id .
                '"></' . CreateElement::NAMESPACEWORD . ':bookmarkEnd>', $this->_xml
        );
    }

    /**
     * Generate w:bookmarkstart
     *
     * @access protected
     * @param int $id
     * @param string $name
     */
    protected function generateBOOKMARKSTART($id, $name)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATER__', '<' . CreateElement::NAMESPACEWORD .
                ':bookmarkStart ' . CreateElement::NAMESPACEWORD . ':id="' . $id .
                '" ' . CreateElement::NAMESPACEWORD . ':name="' . $name . '"></' .
                CreateElement::NAMESPACEWORD .
                ':bookmarkStart>__PHX=__GENERATER____PHX=__GENERATEBOOKMARKEND__', $this->_xml
        );
    }

    /**
     * Generate w:bdr
     *
     * @access protected
     * @param array $border
     */
    protected function generateCHARACTERBORDER($border)
    {
        if (!isset($border['color'])) {
            $border['color'] = 'auto';
        }
        if (!isset($border['spacing'])) {
            $border['spacing'] = 0;
        }
        if (!isset($border['type'])) {
            $border['type'] = 'single';
        }
        if (!isset($border['width'])) {
            $border['width'] = 4;
        }

        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':bdr ' .
                'w:color="' . $border['color'] . '" w:space="' . $border['spacing'] . '" ' .
                'w:sz="' . $border['width'] . '" w:val="' . $border['type'] . '"' .
                '/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:color
     *
     * @access protected
     * @param string $val
     */
    protected function generateCOLOR($val = '000000')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':color ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':color>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:contextualSpacing
     *
     * @access protected
     * @param string $val
     */
    protected function generateCONTEXTUALSPACING($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':contextualSpacing w:val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':contextualSpacing>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:caps
     *
     * @access protected
     * @param string $val
     */
    protected function generateCAPS($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':caps ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':caps>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:dstrike
     *
     * @access protected
     */
    protected function generateDOUBLESTRIKETHROUGH()
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':dstrike/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:em
     *
     * @access protected
     * @param string $val
     */
    protected function generateEM($val = 'none')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':em ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':em>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:emboss
     *
     * @access protected
     * @param string $val
     */
    protected function generateEMBOSS($val = 'on')
    {
        $this->_xml = str_replace(
            '__PHX=__GENERATERPR__',
            '<w:emboss w:val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:lang
     *
     * @access protected
     * @param array $value
     */
    protected function generateLANG($value)
    {
        $attrType = array_keys($value);
        $attrVal = array_values($value);

        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':lang ' .
                CreateElement::NAMESPACEWORD . ':'.$attrType[0].'="' . $attrVal[0] . '"></' .
                CreateElement::NAMESPACEWORD . ':lang>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:outlineLvl
     *
     * @access protected
     * @param string $val
     */
    protected function generateHEADINGLEVEL($val)
    {
        if (is_integer($val) && $val > 0) {
            $this->_xml = str_replace(
                    '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                    ':outlineLvl w:val="' . $val . '"></' .
                    CreateElement::NAMESPACEWORD . ':outlineLvl>__PHX=__GENERATEPPR__', $this->_xml
            );
        }
    }

    /**
     * Generate w:highlight
     *
     * @access protected
     * @param string $val
     */
    protected function generateHIGHLIGHT($val)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':highlight ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':highlight>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:i and w:iCs
     *
     * @access protected
     * @param string $val
     */
    protected function generateI($val = 'single')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':i ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':i><' . CreateElement::NAMESPACEWORD . ':iCs ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':iCs>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:ind
     *
     * @access protected
     * @param array $indentation
     */
    protected function generateINDENT($indentation = array())
    {
        $xmlInd = '<' . CreateElement::NAMESPACEWORD . ':ind ';
        foreach ($indentation as $key => $value) {
            $xmlInd .= CreateElement::NAMESPACEWORD . ':' . $key . '="' . $value . '" ';
        }
        $xmlInd .= '></' . CreateElement::NAMESPACEWORD . ':ind>__PHX=__GENERATEPPR__';
        $this->_xml = str_replace('__PHX=__GENERATEPPR__', $xmlInd, $this->_xml);
    }

    /**
     * Generate w:keepLines
     *
     * @access protected
     * @param string $val
     */
    protected function generateKEEPLINES($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':keepLines w:val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':keepLines>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:keepNext
     *
     * @access protected
     * @param string $val
     */
    protected function generateKEEPNEXT($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':keepNext w:val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':keepNext>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:jc
     *
     * @access protected
     * @param string $val
     */
    protected function generateJC($val = '')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD . ':jc ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':jc>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:noProof
     *
     * @access protected
     * @param string $val
     */
    protected function generateNOPROOF($val = 'on')
    {
        $this->_xml = str_replace(
            '__PHX=__GENERATERPR__',
            '<w:noProof w:val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:outline
     *
     * @access protected
     * @param string $val
     */
    protected function generateOUTLINE($val = 'on')
    {
        $this->_xml = str_replace(
            '__PHX=__GENERATERPR__',
            '<w:outline w:val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:pagebreakbefore
     *
     * @access protected
     * @param string $val
     */
    protected function generatePAGEBREAKBEFORE($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':pageBreakBefore val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':pageBreakBefore>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:tcBorders
     *
     * @param array $border
     * @access protected
     */
    protected function generatePBDR($border)
    {
        $xml = '<w:pBdr>';
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
        $xml .= '</w:pBdr>__PHX=__GENERATEPPR__';
        $this->_xml = str_replace('__PHX=__GENERATEPPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:position
     *
     * @access protected
     * @param int $val
     */
    protected function generatePOSITION($val = 1)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':position ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" />__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:ppr
     *
     * @access protected
     */
    protected function generatePPR()
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':pPr>__PHX=__GENERATEPPR__</' . CreateElement::NAMESPACEWORD .
                ':pPr>__PHX=__GENERATER__';

        $this->_xml = str_replace('__PHX=__GENERATEP__', $xml, $this->_xml);
    }

    /**
     * Generate w:rPr within a w:pPr tag
     *
     * @access protected
     */
    protected function generatePPRRPR()
    {
        /* $xml = '<' . CreateElement::NAMESPACEWORD .
          ':rPr>__PHX=__GENERATERPR__</' . CreateElement::NAMESPACEWORD .
          ':rPr>__PHX=__GENERATER__';

          $this->_xml = str_replace('__PHX=__GENERATER__', $xml, $this->_xml); */

        $this->_xml = str_replace(
            '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD . ':rPr >__PHX=__GENERATERPR__</' .
            CreateElement::NAMESPACEWORD . ':rPr>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:shd
     *
     * @access protected
     * @param string $val
     */
    protected function generatePPRSHD($val = 'FFFFFF')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD . ':shd ' .
                CreateElement::NAMESPACEWORD . ':val="clear" ' .
                CreateElement::NAMESPACEWORD . ':fill="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':shd>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:pstyle
     *
     * @access protected
     * @param string $val
     */
    protected function generatePSTYLE($val = 'TitlePHPDOCX')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD . ':pStyle ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':pStyle>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:rtl
     *
     * @access protected
     * @param string $val
     */
    protected function generateRTL($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':rtl ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':rtl>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:w
     *
     * @access protected
     * @param int $val
     */
    protected function generateSCALING($val = 100)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':w ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:spacing
     *
     * @access protected
     * @param int $val
     */
    protected function generateCHARACTERSPACING($val = 10)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':spacing ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:shadow
     *
     * @access protected
     * @param string $val
     */
    protected function generateSHADOW($val = 'on')
    {
        $this->_xml = str_replace(
            '__PHX=__GENERATERPR__',
            '<w:shadow w:val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:shd
     *
     * @access protected
     * @param string $val
     */
    protected function generateSHD($val = 'FFFFFF')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':shd ' .
                CreateElement::NAMESPACEWORD . ':val="clear" ' .
                CreateElement::NAMESPACEWORD . ':fill="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':shd>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:smallCaps
     *
     * @access protected
     * @param string $val
     */
    protected function generateSMALLCAPS($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':smallCaps ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':smallCaps>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:spacing
     *
     * @access protected
     */
    protected function generateSPACING($spacingTop, $spacingBottom, $line = '240')
    {
        $xml = '<' . CreateElement::NAMESPACEWORD . ':spacing ';
        if (isset($spacingTop) && $spacingTop !== '') {
            $xml .= CreateElement::NAMESPACEWORD . ':before="' . (int) $spacingTop . '" ';
        }
        if (isset($spacingBottom) && $spacingBottom !== '') {
            $xml .= CreateElement::NAMESPACEWORD . ':after="' . (int) $spacingBottom . '" ';
        }
        $xml .= CreateElement::NAMESPACEWORD . ':line="' . $line .
                '" ' . CreateElement::NAMESPACEWORD . ':lineRule="auto"/>__PHX=__GENERATEPPR__';

        $this->_xml = str_replace('__PHX=__GENERATEPPR__', $xml, $this->_xml);
    }

    /**
     * Generate w:strikeThrough
     *
     * @access protected
     */
    protected function generateSTRIKETHROUGH()
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':strike/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:suppressLineNumbers
     *
     * @access protected
     * @param string $val
     */
    protected function generateSUPPRESSLINENUMBERS($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':suppressLineNumbers w:val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':suppressLineNumbers>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:sz and w:szCs
     *
     * @access protected
     * @param string $val
     */
    protected function generateSZ($val = '11')
    {
        $val *= 2;
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':sz ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':sz>' .
                '<' . CreateElement::NAMESPACEWORD . ':szCs ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':szCs>' .
                '__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:t
     *
     * @access protected
     * @param string $dat
     * @param int $spaces
     * @param mixed $lineBreak
     * @param mixed $columnBreak
     */
    protected function generateT($dat, $spaces = 0, $lineBreak = false, $columnBreak = false)
    {
        $strSpaces = '';
        if ($spaces) {
            $i = 0;
            while ($i < $spaces) {
                $strSpaces .= ' ';
                $i++;
            }
        }
        if (is_null($dat)) {
            $dat = '';
        }
        if ($lineBreak == 'before') {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<w:br /><' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t>', $this->_xml
            );
        } else if ($lineBreak == 'after') {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t><w:br />', $this->_xml
            );
        } else if ($lineBreak == 'both') {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<w:br /><' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t><w:br />', $this->_xml
            );
        } else if ($columnBreak == 'before') {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<w:br w:type="column" /><' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t>', $this->_xml
            );
        } else if ($columnBreak == 'after') {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t><w:br w:type="column" />', $this->_xml
            );
        } else if ($columnBreak == 'both') {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<w:br w:type="column" /><' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t><w:br w:type="column" />', $this->_xml
            );
        } else {
            $this->_xml = str_replace(
                    '__PHX=__GENERATER__', '<' . CreateElement::NAMESPACEWORD .
                    ':t xml:space="preserve">' . $strSpaces . $this->parseAndCleanTextString($dat) . '</' .
                    CreateElement::NAMESPACEWORD . ':t>', $this->_xml
            );
        }
    }

    /**
     * Generate w:abs
     *
     * @access protected
     * @param array $tabs
     */
    protected function generateTABPOSITIONS($tabs)
    {
        $typesArray = array('clear', 'left', 'center', 'right', 'decimal', 'bar', 'num');
        $leaderArray = array('none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot');
        $xml = '<w:tabs>';
        foreach ($tabs as $key => $tab) {
            if (isset($tab['type']) && in_array($tab['type'], $typesArray)) {
                $type = $tab['type'];
            } else {
                $type = 'left';
            }
            if (isset($tab['leader']) && in_array($tab['leader'], $leaderArray)) {
                $leader = $tab['leader'];
            } else {
                $leader = 'none';
            }
            if (isset($tab['position']) && is_int($tab['position'])) {
                $xml .='<w:tab w:val="' . $type . '" w:leader="' . $leader . '" w:pos="' . (int) $tab['position'] . '" />';
            }
        }
        $xml .='</w:tabs>';
        $this->_xml = str_replace('__PHX=__GENERATEPPR__', $xml . '__PHX=__GENERATEPPR__', $this->_xml);
    }

    /**
     * Generate w:abs
     *
     * @access protected
     */
    protected function generateTABS()
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATER__', '<' . CreateElement::NAMESPACEWORD .
                ':tab/>__PHX=__GENERATER__', $this->_xml
        );
    }

    /**
     * Generate w:textDirection
     *
     * @access protected
     * @param string $val
     */
    protected function generateTEXTDIRECTION($val = 'lrTb')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':textDirection w:val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':textDirection>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:u
     *
     * @access protected
     * @param string $val
     * @param string $color
     */
    protected function generateU($val = 'single', $color = null)
    {
        if ($color === null) {
            $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':u ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':u>__PHX=__GENERATERPR__', $this->_xml
            );
        } else {
            $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':u ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '" ' .
                CreateElement::NAMESPACEWORD . ':color="' . $color . '"></' .
                CreateElement::NAMESPACEWORD . ':u>__PHX=__GENERATERPR__', $this->_xml
            );
        }
    }

    /**
     * Generate w:vanish
     *
     * @access protected
     */
    protected function generateVANISH()
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':vanish/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:vertAlign
     *
     * @access protected
     * @param string $val
     */
    protected function generateVERTALIGN($val)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATERPR__', '<' . CreateElement::NAMESPACEWORD . ':vertAlign ' .
                CreateElement::NAMESPACEWORD . ':val="' . $val . '"/>__PHX=__GENERATERPR__', $this->_xml
        );
    }

    /**
     * Generate w:widowcontrol
     *
     * @access protected
     * @param string $val
     */
    protected function generateWIDOWCONTROL($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':widowControl ' . CreateElement::NAMESPACEWORD . ':val="' . $val . '"></' .
                CreateElement::NAMESPACEWORD . ':widowControl>__PHX=__GENERATEPPR__', $this->_xml
        );
    }

    /**
     * Generate w:wordwrap
     *
     * @access protected
     * @param string $val
     */
    protected function generateWORDWRAP($val = 'on')
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEPPR__', '<' . CreateElement::NAMESPACEWORD .
                ':wordWrap w:val="' . $val . '"></' . CreateElement::NAMESPACEWORD .
                ':wordWrap>__PHX=__GENERATEPPR__', $this->_xml
        );
    }
}