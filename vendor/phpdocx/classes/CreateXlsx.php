<?php

/**
 * Create XLSX
 *
 * @category   Phpdocx
 * @package    elements
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (http://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class CreateXlsx extends CreateElement implements InterfaceXlsx
{
    /**
     *
     * @access private
     * @var DOCXStructure
     */
    private $_zipXlsx;

    /**
     *
     * @access private
     * @var string
     */
    private $_xmlXlTablesContent;

    /**
     *
     * @access private
     * @var string
     */
    private $_xmlXlSharedStringsContent;

    /**
     *
     * @access private
     * @var string
     */
    private $_xmlXlSheetContent;

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

    // empty functions to be overridden

    /**
     * Create excel table
     *
     * @access public
     * @param array $dats
     */
    public function createExcelTable($dats) {}

    /**
     * Create excel shared strings
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSharedStrings($dats) {}

    /**
     * Create excel sheet
     *
     * @access public
     * @param array $dats
     */
    public function createExcelSheet($dats) {}

    /**
     * Create XLSX
     *
     * @access public
     * @return DOCXStructure
     */
    public function createXlsx()
    {
        $args = func_get_args();
        $this->_zipXlsx = new DOCXStructure();

        $this->_xmlXlTablesContent = '';
        $this->_xmlXlSharedStringsContent = '';
        $this->_xmlXlSheetContent = '';

        $this->_zipXlsx->addContent('[Content_Types].xml', ExcelStructureTemplate::$excelStructure['[Content_Types].xml']);
        $this->_zipXlsx->addContent('docProps/core.xml', ExcelStructureTemplate::$excelStructure['docProps/core.xml']);
        $this->_zipXlsx->addContent('docProps/app.xml', ExcelStructureTemplate::$excelStructure['docProps/app.xml']);
        $this->_zipXlsx->addContent('_rels/.rels', ExcelStructureTemplate::$excelStructure['_rels/.rels']);
        $this->_zipXlsx->addContent('xl/_rels/workbook.xml.rels', ExcelStructureTemplate::$excelStructure['xl/_rels/workbook.xml.rels']);
        $this->_zipXlsx->addContent('xl/theme/theme1.xml', ExcelStructureTemplate::$excelStructure['xl/theme/theme1.xml']);
        $this->_zipXlsx->addContent('xl/worksheets/_rels/sheet1.xml.rels', ExcelStructureTemplate::$excelStructure['xl/worksheets/_rels/sheet1.xml.rels']);
        $this->_zipXlsx->addContent('xl/styles.xml', ExcelStructureTemplate::$excelStructure['xl/styles.xml']);
        $this->_zipXlsx->addContent('xl/workbook.xml', ExcelStructureTemplate::$excelStructure['xl/workbook.xml']);
        $this->_zipXlsx->addContent('xl/tables/table1.xml', $this->createExcelTable($args[1]));
        $this->_zipXlsx->addContent('xl/sharedStrings.xml', $this->createExcelSharedStrings($args[1]));
        $this->_zipXlsx->addContent('xl/worksheets/sheet1.xml', $this->createExcelSheet($args[1]));

        return $this->_zipXlsx;
    }

    /**
     * Generate sst
     *
     * @param mixed $num
     * @access protected
     */
    protected function generateSST($num)
    {
        $this->_xml = '<?xml version="1.0" encoding="UTF-8" ' .
                'standalone="yes" ?><sst xmlns="http://schemas.' .
                'openxmlformats.org/spreadsheetml/2006/main" ' .
                'count="' . $num . '" uniqueCount="' . $num .
                '">__PHX=__GENERATESST__</sst>';
    }

    /**
     * Generate si
     * @access protected
     */
    protected function generateSI()
    {
        $xml = '<si>__PHX=__GENERATESI__</si>__PHX=__GENERATESST__';
        $this->_xml = str_replace('__PHX=__GENERATESST__', $xml, $this->_xml);
    }

    /**
     * Generate t
     *
     * @param string $name
     * @param string $space
     * @access protected
     */
    protected function generateT($name, $space = '')
    {
        $xmlAux = '<t';
        if ($space != '') {
            $xmlAux .= ' xml:space="' . $space . '"';
        }
        $xmlAux .= '>' . $this->parseAndCleanTextString($name) . '</t>';
        $this->_xml = str_replace('__PHX=__GENERATESI__', $xmlAux, $this->_xml);
    }

    /**
     * Generate c
     *
     * @param string $r
     * @param mixed $s
     * @param string $t
     * @access protected
     */
    protected function generateC($r, $s, $t = '')
    {
        $xmlAux = '<c r="' . $r . '"';
        if ($s != '') {
            $xmlAux .= ' s="' . $s . '"';
        }
        if ($t != '') {
            $xmlAux .= ' t="' . $t . '"';
        }
        $xmlAux .= '>__PHX=__GENERATEC__</c>__PHX=__GENERATEROW__';
        $this->_xml = str_replace('__PHX=__GENERATEROW__', $xmlAux, $this->_xml);
    }

    /**
     * Generate col
     *
     * @param string $min
     * @param string $max
     * @param string $width
     * @param string $customWidth
     * @access protected
     */
    protected function generateCOL($min = '1', $max = '1', $width = '11.85546875', $customWidth = '1')
    {
        $xml = '<col min="' . $min . '" max="' . $max . '" width="' . $width .
                '" customWidth="' . $customWidth . '"></col>';

        $this->_xml = str_replace('__PHX=__GENERATECOLS__', $xml, $this->_xml);
    }

    /**
     * Generate cols
     *
     * @access protected
     */
    protected function generateCOLS()
    {
        $xml = '<cols>__PHX=__GENERATECOLS__</cols>__PHX=__GENERATEWORKSHEET__';
        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate dimension
     *
     * @param int $sizeX
     * @param int $sizeY
     * @access protected
     */
    protected function generateDIMENSION($sizeX, $sizeY)
    {
        $char = 'A';
        for ($i = 0; $i < $sizeY; $i++) {
            $char++;
        }
        $sizeX += $sizeY;
        $xml = '<dimension ref="A1:' . $char . $sizeX .
                '"></dimension>__PHX=__GENERATEWORKSHEET__';

        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate pagemargins
     *
     * @param string $left
     * @param string $rigth
     * @param string $bottom
     * @param string $top
     * @param string $header
     * @param string $footer
     * @access protected
     */
    protected function generatePAGEMARGINS($left = '0.7', $rigth = '0.7', $bottom = '0.75', $top = '0.75', $header = '0.3', $footer = '0.3')
    {
        $xml = '<pageMargins left="' . $left . '" right="' . $rigth .
                '" top="' . $top . '" bottom="' . $bottom .
                '" header="' . $header . '" footer="' . $footer .
                '"></pageMargins>__PHX=__GENERATEWORKSHEET__';

        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate pagesetup
     *
     * @param string $paperSize
     * @param string $orientation
     * @param string $id
     * @access protected
     */
    protected function generatePAGESETUP($paperSize = '9', $orientation = 'portrait', $id = '1')
    {
        $xml = '<pageSetup paperSize="' . $paperSize .
                '" orientation="' . $orientation . '" r:id="rId' . $id .
                '"></pageSetup>__PHX=__GENERATEWORKSHEET__';

        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate row
     *
     * @param mixed $r
     * @param mixed $spans
     * @access protected
     */
    protected function generateROW($r, $spans)
    {
        $spans = '1:' . ($spans + 1);
        $xml = '<row r="' . $r . '" spans="' . $spans .
                '">__PHX=__GENERATEROW__</row>__PHX=__GENERATESHEETDATA__';

        $this->_xml = str_replace('__PHX=__GENERATESHEETDATA__', $xml, $this->_xml);
    }

    /**
     * Generate selection
     *
     * @param mixed $num
     * @access protected
     */
    protected function generateSELECTION($num)
    {
        $xml = '<selection activeCell="B' . $num .
                '" sqref="B' . $num . '"></selection>';

        $this->_xml = str_replace('__PHX=__GENERATESHEETVIEW__', $xml, $this->_xml);
    }

    /**
     * Generate sheetdata
     *
     * @access protected
     */
    protected function generateSHEETDATA()
    {
        $xml = '<sheetData>__PHX=__GENERATESHEETDATA__</sheetData>' .
                '__PHX=__GENERATEWORKSHEET__';
        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate sheetformatpr
     *
     * @param string $baseColWidth
     * @param string $defaultRowHeight
     * @access protected
     */
    protected function generateSHEETFORMATPR($baseColWidth = '10', $defaultRowHeight = '15')
    {
        $xml = '<sheetFormatPr baseColWidth="' . $baseColWidth .
                '" defaultRowHeight="' . $defaultRowHeight .
                '"></sheetFormatPr>__PHX=__GENERATEWORKSHEET__';

        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate sheetview
     *
     * @param string $tabSelected
     * @param string $workbookViewId
     * @access protected
     */
    protected function generateSHEETVIEW($tabSelected = '1', $workbookViewId = '0')
    {
        $xml = '<sheetView tabSelected="' . $tabSelected .
                '" workbookViewId="' . $workbookViewId .
                '">__PHX=__GENERATESHEETVIEW__</sheetView>';

        $this->_xml = str_replace('__PHX=__GENERATESHEETVIEWS__', $xml, $this->_xml);
    }

    /**
     * Generate sheetviews
     *
     * @access protected
     */
    protected function generateSHEETVIEWS()
    {
        $xml = '<sheetViews>__PHX=__GENERATESHEETVIEWS__' .
                '</sheetViews>__PHX=__GENERATEWORKSHEET__';
        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate tablepart
     *
     * @param mixed $id
     * @access protected
     */
    protected function generateTABLEPART($id = '1')
    {
        $xml = '<tablePart r:id="rId' . $id . '"></tablePart>';
        $this->_xml = str_replace('__PHX=__GENERATETABLEPARTS__', $xml, $this->_xml);
    }

    /**
     * Generate tableparts
     *
     * @param string $count
     * @access protected
     */
    protected function generateTABLEPARTS($count = '1')
    {
        $xml = '<tableParts count="' . $count .
                '">__PHX=__GENERATETABLEPARTS__</tableParts>';

        $this->_xml = str_replace('__PHX=__GENERATEWORKSHEET__', $xml, $this->_xml);
    }

    /**
     * Generate v
     *
     * @param mixed $num
     * @access protected
     */
    protected function generateV($num)
    {
        $this->_xml = str_replace(
                '__PHX=__GENERATEC__', '<v>' . $num . '</v>', $this->_xml
        );
    }

    /**
     * Generate worksheet
     *
     * @access protected
     */
    protected function generateWORKSHEET()
    {
        $this->_xml = '<?xml version="1.0" encoding="UTF-8" ' .
                'standalone="yes" ?><worksheet ' .
                'xmlns="http://schemas.openxmlformats.org/' .
                'spreadsheetml/2006/main" ' . 'xmlns:r="http://schemas.' .
                'openxmlformats.org/officeDocument/2006/relationships"' .
                '>__PHX=__GENERATEWORKSHEET__</worksheet>';
    }

    /**
     * Clean template row tags
     *
     * @access private
     */
    protected function cleanTemplateROW()
    {
        $this->_xml = str_replace('__PHX=__GENERATEROW__', '', $this->_xml);
    }

    /**
     * Generate table
     *
     * @param int $rows
     * @param int $cols
     * @access protected
     */
    protected function generateTABLE($rows, $cols)
    {
        $word = 'A';
        for ($i = 0; $i < $cols; $i++) {
            $word++;
        }
        $rows++;
        $this->_xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
                '<table xmlns="http://schemas.openxmlformats.org/spreads' .
                'heetml/2006/main" id="1" name="Tabla1" displayName=' .
                '"Tabla1" ref="A1:' . $word . $rows .
                '" totalsRowShown="0" tableBorderDxfId="0">' .
                '__PHX=__GENERATETABLE__</table>';
    }

    /**
     * Generate tablecolumn
     *
     * @param mixed $id
     * @param string $name
     * @access protected
     */
    protected function generateTABLECOLUMN($id = '2', $name = '')
    {
        $xml = '<tableColumn id="' . $id . '" name="' . $name .
                '"></tableColumn >__PHX=__GENERATETABLECOLUMNS__';

        $this->_xml = str_replace(
                '__PHX=__GENERATETABLECOLUMNS__', $xml, $this->_xml
        );
    }

    /**
     * Generate tablecolumns
     *
     * @param mixed $count
     * @access protected
     */
    protected function generateTABLECOLUMNS($count = '2')
    {
        $xml = '<tableColumns count="' . $count .
                '">__PHX=__GENERATETABLECOLUMNS__</tableColumns>__PHX=__GENERATETABLE__';

        $this->_xml = str_replace('__PHX=__GENERATETABLE__', $xml, $this->_xml);
    }

    /**
     * Generate tablestyleinfo
     *
     * @param string $showFirstColumn
     * @param string $showLastColumn
     * @param string $showRowStripes
     * @param string $showColumnStripes
     * @access protected
     */
    protected function generateTABLESTYLEINFO($showFirstColumn = '0', $showLastColumn = "0", $showRowStripes = "1", $showColumnStripes = "0")
    {
        $xml = '<tableStyleInfo   showFirstColumn="' . $showFirstColumn .
                '" showLastColumn="' . $showLastColumn .
                '" showRowStripes="' . $showRowStripes .
                '" showColumnStripes="' . $showColumnStripes .
                '"></tableStyleInfo >';

        $this->_xml = str_replace('__PHX=__GENERATETABLE__', $xml, $this->_xml);
    }
}