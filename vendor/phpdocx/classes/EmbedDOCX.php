<?php

/**
 * Create DOCX from DOCX file
 *
 * @category   Phpdocx
 * @package    elements
 * @package    transform
 * @copyright  Copyright (c) Narcea Producciones Multimedia S.L.
 *             (https://www.2mdc.com)
 * @license    phpdocx LICENSE
 * @link       https://www.phpdocx.com
 */
class EmbedDOCX extends CreateElement implements EmbedDocument
{
    /**
     *
     * @access private
     * @var mixed
     */
    private $_id = 0;

    /**
     *
     * @access private
     * @var mixed
     */
    private static $_instance = NULL;

    /**
     *
     * @access protected
     * @var string
     */
    protected $_xml = '';

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
     * @return EmbedDOCX
     * @static
     */
    public static function getInstance()
    {
        if (self::$_instance == NULL) {
            self::$_instance = new EmbedDOCX();
        }
        return self::$_instance;
    }

    /**
     * Getter. Return current HTML ID
     *
     * @access public
     * @return mixed
     */
    public function getId()
    {
        return $this->_id;
    }

    /**
     * Embed HTML in DOCX
     *
     * @param bool $matchSource
     * @access public
     */
    public function embed($matchSource = false)
    {
        $this->_xml = '';
        $this->_id = uniqid((string)mt_rand(999, 9999));
        $this->generateALTCHUNK($matchSource);
    }

    /**
     * Generate w:altChunk
     *
     * @param bool $matchSource
     * @access protected
     */
    public function generateALTCHUNK($matchSource = false)
    {
        $this->_xml = '<' . CreateElement::NAMESPACEWORD .
                ':altChunk r:id="rDOCXId' . $this->_id . '" ' .
                'xmlns:r="http://schemas.openxmlformats.org/' .
                'officeDocument/2006/relationships" ' .
                'xmlns:w="http://schemas.openxmlformats.org/' .
                'wordprocessingml/2006/main" >';
        if ($matchSource) {
            $this->_xml .= '<w:altChunkPr><w:matchSrc/></w:altChunkPr>';
        }

        $this->_xml .= '</w:altChunk>';
    }

}