<?php
/**
 * This file is part of PHPWord - A pure PHP library for reading and writing
 * word processing documents.
 *
 * PHPWord is free software distributed under the terms of the GNU Lesser
 * General Public License version 3 as published by the Free Software Foundation.
 *
 * For the full copyright and license information, please read the LICENSE
 * file that was distributed with this source code. For the full list of
 * contributors, visit https://github.com/PHPOffice/PHPWord/contributors.
 *
 * @link        https://github.com/PHPOffice/PHPWord
 * @copyright   2010-2016 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord;

use PhpOffice\PhpWord\Escaper\RegExp;
use PhpOffice\PhpWord\Escaper\Xml;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\Shared\ZipArchive;
use Zend\Stdlib\StringUtils;

class TemplateProcessor
{
    const MAXIMUM_REPLACEMENTS_DEFAULT = -1;

    /**
     * ZipArchive object.
     *
     * @var mixed
     */
    protected $zipClass;

    /**
     * @var string Temporary document filename (with path).
     */
    protected $tempDocumentFilename;

    /**
     * Content of main document part (in XML format) of the temporary document.
     *
     * @var string
     */
    protected $tempDocumentMainPart;

    /**
     * Content of headers (in XML format) of the temporary document.
     *
     * @var string[]
     */
    protected $tempDocumentHeaders = array();

    /**
     * Content of footers (in XML format) of the temporary document.
     *
     * @var string[]
     */
    protected $tempDocumentFooters = array();

    /**
     * Content of main rels document part (in XML format) of the temporary document.
     *
     * @var string
     */
    protected $tempRelsDocumentMainPart;

    /**
     * Content of rels headers (in XML format) of the temporary document.
     *
     * @var string[]
     */
    protected $tempRelsDocumentHeaders = array();

    /**
     * Content of rels footers (in XML format) of the temporary document.
     *
     * @var string[]
     */
    protected $tempRelsDocumentFooters = array();

    /**
     * @since 0.12.0 Throws CreateTemporaryFileException and CopyFileException instead of Exception.
     *
     * @param string $documentTemplate The fully qualified template filename.
     *
     * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
     * @throws \PhpOffice\PhpWord\Exception\CopyFileException
     */
    public function __construct($documentTemplate)
    {
        // Temporary document filename initialization
        $this->tempDocumentFilename = tempnam(Settings::getTempDir(), 'PhpWord');
        if (false === $this->tempDocumentFilename) {
            throw new CreateTemporaryFileException();
        }

        // Template file cloning
        if (false === copy($documentTemplate, $this->tempDocumentFilename)) {
            throw new CopyFileException($documentTemplate, $this->tempDocumentFilename);
        }

        // Temporary document content extraction
        $this->zipClass = new ZipArchive();
        $this->zipClass->open($this->tempDocumentFilename);
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getHeaderName($index))) {
            $this->tempDocumentHeaders[$index] = $this->fixBrokenMacros(
                $this->zipClass->getFromName($this->getHeaderName($index))
            );
            $index++;
        }
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getFooterName($index))) {
            $this->tempDocumentFooters[$index] = $this->fixBrokenMacros(
                $this->zipClass->getFromName($this->getFooterName($index))
            );
            $index++;
        }
        $this->tempDocumentMainPart = $this->fixBrokenMacros($this->zipClass->getFromName($this->getMainPartName()));

        $index = 1;
        while (false !== $this->zipClass->locateName($this->getRelsHeaderName($index))) {
            $this->tempRelsDocumentHeaders[$index] = $this->fixBrokenMacros(
                $this->zipClass->getFromName($this->getRelsHeaderName($index))
            );
            $index++;
        }
        $index = 1;
        while (false !== $this->zipClass->locateName($this->getRelsFooterName($index))) {
            $this->tempRelsDocumentFooters[$index] = $this->fixBrokenMacros(
                $this->zipClass->getFromName($this->getRelsFooterName($index))
            );
            $index++;
        }
        $this->tempRelsDocumentMainPart = $this->fixBrokenMacros($this->zipClass->getFromName($this->getRelsMainPartName()));
    }

    /**
     * @param string $xml
     * @param \XSLTProcessor $xsltProcessor
     *
     * @return string
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    protected function transformSingleXml($xml, $xsltProcessor)
    {
        $domDocument = new \DOMDocument();
        if (false === $domDocument->loadXML($xml)) {
            throw new Exception('Could not load the given XML document.');
        }

        $transformedXml = $xsltProcessor->transformToXml($domDocument);
        if (false === $transformedXml) {
            throw new Exception('Could not transform the given XML document.');
        }

        return $transformedXml;
    }

    /**
     * @param mixed $xml
     * @param \XSLTProcessor $xsltProcessor
     *
     * @return mixed
     */
    protected function transformXml($xml, $xsltProcessor)
    {
        if (is_array($xml)) {
            foreach ($xml as &$item) {
                $item = $this->transformSingleXml($item, $xsltProcessor);
            }
        } else {
            $xml = $this->transformSingleXml($xml, $xsltProcessor);
        }

        return $xml;
    }

    /**
     * Applies XSL style sheet to template's parts.
     *
     * Note: since the method doesn't make any guess on logic of the provided XSL style sheet,
     * make sure that output is correctly escaped. Otherwise you may get broken document.
     *
     * @param \DOMDocument $xslDomDocument
     * @param array $xslOptions
     * @param string $xslOptionsUri
     *
     * @return void
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function applyXslStyleSheet($xslDomDocument, $xslOptions = array(), $xslOptionsUri = '')
    {
        $xsltProcessor = new \XSLTProcessor();

        $xsltProcessor->importStylesheet($xslDomDocument);
        if (false === $xsltProcessor->setParameter($xslOptionsUri, $xslOptions)) {
            throw new Exception('Could not set values for the given XSL style sheet parameters.');
        }

        $this->tempDocumentHeaders = $this->transformXml($this->tempDocumentHeaders, $xsltProcessor);
        $this->tempDocumentMainPart = $this->transformXml($this->tempDocumentMainPart, $xsltProcessor);
        $this->tempDocumentFooters = $this->transformXml($this->tempDocumentFooters, $xsltProcessor);
    }

    /**
     * @param string $macro
     *
     * @return string
     */
    protected static function ensureMacroCompleted($macro)
    {
        if (substr($macro, 0, 2) !== '${' && substr($macro, -1) !== '}') {
            $macro = '${' . $macro . '}';
        }

        return $macro;
    }

    /**
     * @param string $subject
     *
     * @return string
     */
    protected static function ensureUtf8Encoded($subject)
    {
        if (!StringUtils::isValidUtf8($subject)) {
            $subject = utf8_encode($subject);
        }

        return $subject;
    }

    /**
     * @param mixed $search
     * @param mixed $replace
     * @param integer $limit
     *
     * @return void
     */
    public function setValue($search, $replace, $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT)
    {
        $replace = str_replace(
            ['&lt;', '&gt;', '&'],
            ['_lt_', '_gt_', '_amp_'],
            $replace
        );
        $replace = preg_replace('~\R~u', '</w:t><w:br/><w:t>', $replace);

        if (is_array($search)) {
            foreach ($search as &$item) {
                $item = self::ensureMacroCompleted($item);
            }
        } else {
            $search = self::ensureMacroCompleted($search);
        }

        if (is_array($replace)) {
            foreach ($replace as &$item) {
                $item = self::ensureUtf8Encoded($item);
            }
        } else {
            $replace = self::ensureUtf8Encoded($replace);
        }

        if (Settings::isOutputEscapingEnabled()) {
            $xmlEscaper = new Xml();
            $replace = $xmlEscaper->escape($replace);
        }

        $this->tempDocumentHeaders = $this->setValueForPart($search, $replace, $this->tempDocumentHeaders, $limit);
        $this->tempDocumentMainPart = $this->setValueForPart($search, $replace, $this->tempDocumentMainPart, $limit);
        $this->tempDocumentFooters = $this->setValueForPart($search, $replace, $this->tempDocumentFooters, $limit);
    }

    /**
     * @param string $search
     * @param string $replace
     * @param integer $limit
     *
     * @return void
     */
    public function setHtml($search, $replace, $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT)
    {
        $search = self::ensureMacroCompleted($search);
        //$replace = self::ensureUtf8Encoded($replace);

        $replace = str_replace(
            ['<br><br>','<br>', '<div>', '</div>', '&lt;', '&gt;', '&amp;'],
            ['<br/>','<br/>', '', '', '', '_gt_', ''],
            $replace
        );

        // Turn it into word data
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        \PhpOffice\PhpWord\Shared\Html::addHtml($section, $replace);

        $part = new \PhpOffice\PhpWord\Writer\Word2007\Part\Document();
        $part->setParentWriter(new \PhpOffice\PhpWord\Writer\Word2007($phpWord));
        $replace = $part->write();

        $this->tempDocumentMainPart = $this->setValueForPartHtml($search, $replace, $this->tempDocumentMainPart, $limit);

        foreach($this->tempDocumentHeaders as $i => $h){
            $this->tempDocumentHeaders[$i] = $this->setValueForPartHtml($search, $replace, $this->tempDocumentHeaders[$i], $limit);
        }

        foreach($this->tempDocumentFooters as $i => $h){
            $this->tempDocumentFooters[$i] = $this->setValueForPartHtml($search, $replace, $this->tempDocumentFooters[$i], $limit);
        }
    }

    protected function setValueForPartHtml($search, $replace, $documentPartXML, $limit)
    {
        if(!empty($documentPartXML)){
            $regExpDelim = '/';
            $escapedSearch = preg_quote($search, $regExpDelim);
            $xml = new \SimpleXMLElement($documentPartXML);
            foreach ($xml->xpath("//w:p/*[contains(.,'{$search}')]/parent::*") as $node) {
                $count = 0;
                $escapedSearch = preg_quote($node->asXML(), $regExpDelim);

                $attr = $node->attributes();
                try {
                    $xmlReplace = new \SimpleXMLElement($replace);
                } catch (\Exception $e) {
                    continue;
                }
                foreach($xmlReplace->xpath("//w:p") as &$sub){
                    foreach($attr as $a => $b){
                        $sub->addAttribute($a,$b);
                    }
                }

                $replaceData = $xmlReplace->asXML();
                if (preg_match('/<w:body>(.*)<w:sectPr>/is', $replaceData, $matches) === 1) {
                    $replaceData = $matches[1];
                }
                $replaceData = str_replace(['w:val=""'],['w:val="1"'],$replaceData);

                $documentPartXML = preg_replace("{$regExpDelim}{$escapedSearch}{$regExpDelim}u", $replaceData, $documentPartXML, $limit,$count);
                if($limit != self::MAXIMUM_REPLACEMENTS_DEFAULT){
                    $limit -= $count;
                }
                if($limit == 0){
                    break;
                }
            }
        }
        return $documentPartXML;
    }

    /**
     * @param string $search
     * @param string $path
     * @param string[] $options
     * @param integer $limit
     *
     * @return void
     */
    public function setImg($search, $path, $options = array(), $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT){
        $search = self::ensureMacroCompleted($search);
        if(!file_exists($path)){
            return;
        }

        list($this->tempDocumentMainPart, $this->tempRelsDocumentMainPart) = $this->setImgForPart($this->tempDocumentMainPart, $this->tempRelsDocumentMainPart, $search, $path, $options, $limit);

        foreach($this->tempDocumentHeaders as $i => $h){
            list($this->tempDocumentHeaders[$i], $this->tempRelsDocumentHeaders[$i]) = $this->setImgForPart($h, isset($this->tempRelsDocumentHeaders[$i])?$this->tempRelsDocumentHeaders[$i]:'', $search, $path, $options, $limit);
        }
        foreach($this->tempDocumentFooters as $i => $f){
            list($this->tempDocumentFooters[$i], $this->tempRelsDocumentFooters[$i]) = $this->setImgForPart($f, isset($this->tempRelsDocumentFooters[$i])?$this->tempRelsDocumentFooters[$i]:'', $search, $path, $options, $limit);
        }
    }

    protected function setImgForPart($document, $relDocument, $search, $path, $options = array(), $limit = self::MAXIMUM_REPLACEMENTS_DEFAULT){
        if(strpos($document, $search) !== false){
            $id = $filename = null;
            if(empty($relDocument)){
                $relDocument = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
            }
            $path_parts = pathinfo($path);
            $filename = $path_parts['basename'];
            $domDocument = new \DOMDocument();
            if (false === $domDocument->loadXML($relDocument)) {
                throw new Exception('Could not load the given XML document.');
            }
            $id = $domDocument->getElementsByTagName('Relationship')->length +1;
            $this->zipClass->addFile($path,'word/media/'.$filename);

            // Build Relationship datas
            $relationShip = $domDocument->createElement("Relationship");
            $att = $domDocument->createAttribute('Id');
            $att->value = "rId".$id;
            $relationShip->appendChild($att);
            $att = $domDocument->createAttribute('Target');
            $att->value = 'media/'.$filename;
            $relationShip->appendChild($att);
            $att = $domDocument->createAttribute('Type');
            $att->value = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
            $relationShip->appendChild($att);
            $domDocument->documentElement->appendChild($relationShip);
            $relDocument = $domDocument->saveXML();

            if(!empty($id) && !empty($filename)){
                $image = new \PhpOffice\PhpWord\Element\Image($path,$options);
                $image->setRelationId($id);
                $image->setDocPart("");
                // We assume its docx ?
                $writer = new \PhpOffice\Common\XMLWriter();
                $imgWriter = new \PhpOffice\PhpWord\Writer\Word2007\Element\Image($writer, $image);
                $imgWriter->write();
                $replace = trim($writer->getData());
                $document = $this->setValueForPartContent($search, $replace, $document, $limit);
            }
        }
        return [$document, $relDocument];
    }

    /**
     * Returns array of all variables in template.
     *
     * @return string[]
     */
    public function getVariables()
    {
        $variables = $this->getVariablesForPart($this->tempDocumentMainPart);

        foreach ($this->tempDocumentHeaders as $headerXML) {
            $variables = array_merge($variables, $this->getVariablesForPart($headerXML));
        }

        foreach ($this->tempDocumentFooters as $footerXML) {
            $variables = array_merge($variables, $this->getVariablesForPart($footerXML));
        }

        return array_unique($variables);
    }

    /**
     * Clone a table row in a template document.
     *
     * @param string $search
     * @param integer $numberOfClones
     *
     * @return void
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function cloneRow($search, $numberOfClones)
    {
        if ('${' !== substr($search, 0, 2) && '}' !== substr($search, -1)) {
            $search = '${' . $search . '}';
        }

        $tagPos = strpos($this->tempDocumentMainPart, $search);
        if (!$tagPos) {
            throw new Exception("Can not clone row, template variable not found or variable contains markup.");
        }

        $rowStart = $this->findRowStart($tagPos);
        $rowEnd = $this->findRowEnd($tagPos);
        $xmlRow = $this->getSlice($rowStart, $rowEnd);

        // Check if there's a cell spanning multiple rows.
        if (preg_match('#<w:vMerge w:val="restart"/>#', $xmlRow)) {
            // $extraRowStart = $rowEnd;
            $extraRowEnd = $rowEnd;
            while (true) {
                $extraRowStart = $this->findRowStart($extraRowEnd + 1);
                $extraRowEnd = $this->findRowEnd($extraRowEnd + 1);

                // If extraRowEnd is lower then 7, there was no next row found.
                if ($extraRowEnd < 7) {
                    break;
                }

                // If tmpXmlRow doesn't contain continue, this row is no longer part of the spanned row.
                $tmpXmlRow = $this->getSlice($extraRowStart, $extraRowEnd);
                if (!preg_match('#<w:vMerge/>#', $tmpXmlRow) &&
                    !preg_match('#<w:vMerge w:val="continue" />#', $tmpXmlRow) &&
                    !preg_match('#<w:vMerge w:val="continue"/>#', $tmpXmlRow)) {
                    break;
                }
                // This row was a spanned row, update $rowEnd and search for the next row.
                $rowEnd = $extraRowEnd;
            }
            $xmlRow = $this->getSlice($rowStart, $rowEnd);
        }

        $result = $this->getSlice(0, $rowStart);
        for ($i = 1; $i <= $numberOfClones; $i++) {
            $result .= preg_replace('/\$\{(.*?)\}/', '\${\\1#' . $i . '}', $xmlRow);
        }
        $result .= $this->getSlice($rowEnd);

        $this->tempDocumentMainPart = $result;
    }

    /**
     * Clone a block.
     *
     * @param string $blockname
     * @param integer $clones
     * @param boolean $replace
     *
     * @return string|null
     */
    public function cloneBlock($blockname, $clones = 1, $replace = true)
    {
        $xmlBlock = null;
        preg_match(
            '/(<\?xml.*)(<w:p.*>\${' . $blockname . '}<\/w:.*?p>)(.*)(<w:p.*\${\/' . $blockname . '}<\/w:.*?p>)/is',
            $this->tempDocumentMainPart,
            $matches
        );

        if (isset($matches[3])) {
            $xmlBlock = $matches[3];
            $cloned = array();
            for ($i = 1; $i <= $clones; $i++) {
                $cloned[] = $xmlBlock;
            }

            if ($replace) {
                $this->tempDocumentMainPart = str_replace(
                    $matches[2] . $matches[3] . $matches[4],
                    implode('', $cloned),
                    $this->tempDocumentMainPart
                );
            }
        }

        return $xmlBlock;
    }

    /**
     * Replace a block.
     *
     * @param string $blockname
     * @param string $replacement
     *
     * @return void
     */
    public function replaceBlock($blockname, $replacement)
    {
        preg_match(
            '/(<\?xml.*)(<w:p.*>\${' . $blockname . '}<\/w:.*?p>)(.*)(<w:p.*\${\/' . $blockname . '}<\/w:.*?p>)/is',
            $this->tempDocumentMainPart,
            $matches
        );

        if (isset($matches[3])) {
            $this->tempDocumentMainPart = str_replace(
                $matches[2] . $matches[3] . $matches[4],
                $replacement,
                $this->tempDocumentMainPart
            );
        }
    }

    /**
     * Delete a block of text.
     *
     * @param string $blockname
     *
     * @return void
     */
    public function deleteBlock($blockname)
    {
        $this->replaceBlock($blockname, '');
    }

    /**
     * Saves the result document.
     *
     * @return string
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    public function save()
    {
        foreach ($this->tempDocumentHeaders as $index => $xml) {
            $xml = str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $xml);
            $this->zipClass->addFromString($this->getHeaderName($index), $xml);
        }

        $this->zipClass->addFromString($this->getMainPartName(), str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $this->tempDocumentMainPart));

        foreach ($this->tempDocumentFooters as $index => $xml) {
            $xml = str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $xml);
            $this->zipClass->addFromString($this->getFooterName($index), $xml);
        }

        foreach ($this->tempRelsDocumentHeaders as $index => $xml) {
            $xml = str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $xml);
            $this->zipClass->addFromString($this->getRelsHeaderName($index), $xml);
        }

        $this->zipClass->addFromString($this->getRelsMainPartName(), str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $this->tempRelsDocumentMainPart));

        foreach ($this->tempRelsDocumentFooters as $index => $xml) {
            $xml = str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $xml);
            $this->zipClass->addFromString($this->getRelsFooterName($index), $xml);
        }

        // Close zip file
        if (false === $this->zipClass->close()) {
            throw new Exception('Could not close zip file.');
        }

        return $this->tempDocumentFilename;
    }

    /**
     * Saves the result document to the user defined file.
     *
     * @since 0.8.0
     *
     * @param string $fileName
     *
     * @return void
     */
    public function saveAs($fileName)
    {
        $tempFileName = $this->save();

        if (file_exists($fileName)) {
            unlink($fileName);
        }

        /*
         * Note: we do not use `rename` function here, because it looses file ownership data on Windows platform.
         * As a result, user cannot open the file directly getting "Access denied" message.
         *
         * @see https://github.com/PHPOffice/PHPWord/issues/532
         */
        copy($tempFileName, $fileName);
        unlink($tempFileName);
    }

    /**
     * Finds parts of broken macros and sticks them together.
     * Macros, while being edited, could be implicitly broken by some of the word processors.
     *
     * @param string $documentPart The document part in XML representation.
     *
     * @return string
     */
    protected function fixBrokenMacros($documentPart)
    {
        $fixedDocumentPart = $documentPart;

        $fixedDocumentPart = preg_replace_callback(
            '|\$[^{]*\{[^}]*\}|U',
            function ($match) {
                return strip_tags($match[0]);
            },
            $fixedDocumentPart
        );

        return $fixedDocumentPart;
    }

    /**
     * Find and replace macros in the given XML section.
     *
     * @param mixed $search
     * @param mixed $replace
     * @param string $documentPartXML
     * @param integer $limit
     *
     * @return string
     */
    protected function setValueForPart($search, $replace, $documentPartXML, $limit)
    {
        if(strpos($replace, '<w:tbl>') === 0){
            if(is_array($documentPartXML)){
                foreach($documentPartXML as &$doc){
                    $doc = $this->setValueForPartContent($search, $replace, $doc, $limit);
                }
            }else{
                $documentPartXML = $this->setValueForPartContent($search, $replace, $documentPartXML, $limit);
            }
            return $documentPartXML;
        }else{
            // Note: we can't use the same function for both cases here, because of performance considerations.
            if (self::MAXIMUM_REPLACEMENTS_DEFAULT === $limit) {
                return str_replace($search, $replace, $documentPartXML);
            } else {
                $regExpEscaper = new RegExp();
                return preg_replace($regExpEscaper->escape($search), $replace, $documentPartXML, $limit);
            }
        }
    }

    protected function setValueForPartContent($search, $replace, $documentPartXML, $limit)
    {
        if(!empty($documentPartXML)){
            $regExpDelim = '/';
            $escapedSearch = preg_quote($search, $regExpDelim);
            $search = preg_replace("/&amp;/", " ", $search);
            try {
                $xml = new \SimpleXMLElement($documentPartXML);
            } catch (\Exception $e) {
                return $documentPartXML;
            }
            // $xml = new \SimpleXMLElement($documentPartXML);
            foreach ($xml->xpath("//w:p/*[contains(.,'{$search}')]/parent::*") as $node) {
                $count = 0;
                $escapedSearch = preg_quote($node->asXML(), $regExpDelim);
                $documentPartXML = preg_replace("{$regExpDelim}{$escapedSearch}{$regExpDelim}u", $replace, $documentPartXML, $limit,$count);
                if($limit != self::MAXIMUM_REPLACEMENTS_DEFAULT){
                    $limit -= $count;
                }
                if($limit == 0){
                    break;
                }
            }
        }
        return $documentPartXML;
    }

    /**
     * Find all variables in $documentPartXML.
     *
     * @param string $documentPartXML
     *
     * @return string[]
     */
    protected function getVariablesForPart($documentPartXML)
    {
        preg_match_all('/\$\{(.*?)}/i', $documentPartXML, $matches);

        return $matches[1];
    }

    /**
     * Get the name of the header file for $index.
     *
     * @param integer $index
     *
     * @return string
     */
    protected function getHeaderName($index)
    {
        return sprintf('word/header%d.xml', $index);
    }

    /**
     * @return string
     */
    protected function getMainPartName()
    {
        return 'word/document.xml';
    }

    /**
     * Get the name of the footer file for $index.
     *
     * @param integer $index
     *
     * @return string
     */
    protected function getFooterName($index)
    {
        return sprintf('word/footer%d.xml', $index);
    }

    /**
     * Get the name of the rels header file for $index.
     *
     * @param integer $index
     *
     * @return string
     */
    protected function getRelsHeaderName($index)
    {
        return sprintf('word/_rels/header%d.xml.rels', $index);
    }

    /**
     * @return string
     */
    protected function getRelsMainPartName()
    {
        return 'word/_rels/document.xml.rels';
    }

    /**
     * Get the name of the rels footer file for $index.
     *
     * @param integer $index
     *
     * @return string
     */
    protected function getRelsFooterName($index)
    {
        return sprintf('word/_rels/footer%d.xml.rels', $index);
    }

    /**
     * Find the start position of the nearest table row before $offset.
     *
     * @param integer $offset
     *
     * @return integer
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    protected function findRowStart($offset)
    {
        $rowStart = strrpos($this->tempDocumentMainPart, '<w:tr ', ((strlen($this->tempDocumentMainPart) - $offset) * -1));

        if (!$rowStart) {
            $rowStart = strrpos($this->tempDocumentMainPart, '<w:tr>', ((strlen($this->tempDocumentMainPart) - $offset) * -1));
        }
        if (!$rowStart) {
            throw new Exception('Can not find the start position of the row to clone.');
        }

        return $rowStart;
    }

    /**
     * Find the end position of the nearest table row after $offset.
     *
     * @param integer $offset
     *
     * @return integer
     */
    protected function findRowEnd($offset)
    {
        return strpos($this->tempDocumentMainPart, '</w:tr>', $offset) + 7;
    }

    /**
     * Get a slice of a string.
     *
     * @param integer $startPosition
     * @param integer $endPosition
     *
     * @return string
     */
    protected function getSlice($startPosition, $endPosition = 0)
    {
        if (!$endPosition) {
            $endPosition = strlen($this->tempDocumentMainPart);
        }

        return substr($this->tempDocumentMainPart, $startPosition, ($endPosition - $startPosition));
    }
}
