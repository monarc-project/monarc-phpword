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
 * @copyright   2010-2014 PHPWord contributors
 * @license     http://www.gnu.org/licenses/lgpl.txt LGPL version 3
 */

namespace PhpOffice\PhpWord;

use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\Shared\WordString;
use PhpOffice\PhpWord\Shared\ZipArchive;

class TemplateProcessor {
	/**
	 * ZipArchive object.
	 *
	 * @var mixed
	 */
	private $zipClass;

	/**
	 * @var string Temporary document filename (with path).
	 */
	private $temporaryDocumentFilename;

	/**
	 * Content of main document part (in XML format) of the temporary document.
	 *
	 * @var string
	 */
	private $temporaryDocumentMainPart;

	/**
	 * Content of headers (in XML format) of the temporary document.
	 *
	 * @var string[]
	 */
	private $temporaryDocumentHeaders = array();

	/**
	 * Content of footers (in XML format) of the temporary document.
	 *
	 * @var string[]
	 */
	private $temporaryDocumentFooters = array();

	/**
	 * @since 0.12.0 Throws CreateTemporaryFileException and CopyFileException instead of Exception.
	 *
	 * @param string $documentTemplate The fully qualified template filename.
	 * @throws \PhpOffice\PhpWord\Exception\CreateTemporaryFileException
	 * @throws \PhpOffice\PhpWord\Exception\CopyFileException
	 */
	public function __construct($documentTemplate) {
		// Temporary document filename initialization
		$this->temporaryDocumentFilename = tempnam(Settings::getTempDir(), 'PhpWord');
		if (false === $this->temporaryDocumentFilename) {
			throw new CreateTemporaryFileException();
		}

		// Template file cloning
		if (false === copy($documentTemplate, $this->temporaryDocumentFilename)) {
			throw new CopyFileException($documentTemplate, $this->temporaryDocumentFilename);
		}

		// Temporary document content extraction
		$this->zipClass = new ZipArchive();
		$this->zipClass->open($this->temporaryDocumentFilename);
		$index = 1;
		while (false !== $this->zipClass->locateName($this->getHeaderName($index))) {
			$this->temporaryDocumentHeaders[$index] = $this->fixBrokenMacros(
				$this->zipClass->getFromName($this->getHeaderName($index))
			);
			$index++;
		}
		$index = 1;
		while (false !== $this->zipClass->locateName($this->getFooterName($index))) {
			$this->temporaryDocumentFooters[$index] = $this->fixBrokenMacros(
				$this->zipClass->getFromName($this->getFooterName($index))
			);
			$index++;
		}
		$this->temporaryDocumentMainPart = $this->fixBrokenMacros($this->zipClass->getFromName('word/document.xml'));
	}

	/**
	 * Applies XSL style sheet to template's parts.
	 *
	 * @param \DOMDocument $xslDOMDocument
	 * @param array $xslOptions
	 * @param string $xslOptionsURI
	 * @return void
	 * @throws \PhpOffice\PhpWord\Exception\Exception
	 */
	public function applyXslStyleSheet($xslDOMDocument, $xslOptions = array(), $xslOptionsURI = '') {
		$xsltProcessor = new \XSLTProcessor();

		$xsltProcessor->importStylesheet($xslDOMDocument);

		if (false === $xsltProcessor->setParameter($xslOptionsURI, $xslOptions)) {
			throw new Exception('Could not set values for the given XSL style sheet parameters.');
		}

		$xmlDOMDocument = new \DOMDocument();
		if (false === $xmlDOMDocument->loadXML($this->temporaryDocumentMainPart)) {
			throw new Exception('Could not load XML from the given template.');
		}

		$xmlTransformed = $xsltProcessor->transformToXml($xmlDOMDocument);
		if (false === $xmlTransformed) {
			throw new Exception('Could not transform the given XML document.');
		}

		$this->temporaryDocumentMainPart = $xmlTransformed;
	}

	/**
	 * @param mixed $search
	 * @param mixed $replace
	 * @param integer $limit
	 * @return void
	 */
	public function setValue($search, $replace, $limit = -1) {
		$replace = str_replace('&', '&amp;', $replace);
		$replace = preg_replace('~\R~u', '</w:t><w:br/><w:t>', $replace);

		foreach ($this->temporaryDocumentHeaders as $index => $headerXML) {
			$this->temporaryDocumentHeaders[$index] = $this->setValueForPart($this->temporaryDocumentHeaders[$index], $search, $replace, $limit);
		}

		$this->temporaryDocumentMainPart = $this->setValueForPart($this->temporaryDocumentMainPart, $search, $replace, $limit);

		foreach ($this->temporaryDocumentFooters as $index => $headerXML) {
			$this->temporaryDocumentFooters[$index] = $this->setValueForPart($this->temporaryDocumentFooters[$index], $search, $replace, $limit);
		}
	}

	/**
	 * Returns array of all variables in template.
	 *
	 * @return string[]
	 */
	public function getVariables() {
		$variables = $this->getVariablesForPart($this->temporaryDocumentMainPart);

		foreach ($this->temporaryDocumentHeaders as $headerXML) {
			$variables = array_merge($variables, $this->getVariablesForPart($headerXML));
		}

		foreach ($this->temporaryDocumentFooters as $footerXML) {
			$variables = array_merge($variables, $this->getVariablesForPart($footerXML));
		}

		return array_unique($variables);
	}

	/**
	 * Clone a table row in a template document.
	 *
	 * @param string $search
	 * @param integer $numberOfClones
	 * @return void
	 * @throws \PhpOffice\PhpWord\Exception\Exception
	 */
	public function cloneRow($search, $numberOfClones) {
		if ('${' !== substr($search, 0, 2) && '}' !== substr($search, -1)) {
			$search = '${' . $search . '}';
		}

		$tagPos = strpos($this->temporaryDocumentMainPart, $search);
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
					!preg_match('#<w:vMerge w:val="continue"/>#', $tmpXmlRow)
				) {
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

		$this->temporaryDocumentMainPart = $result;
	}

	/**
	 * Clone a block
	 *
	 * @param string $blockname
	 * @param integer $clones
	 * @param boolean $replace
	 * @return string|null
	 */
	public function cloneBlock($blockname, $clones = 1, $replace = true) {
		// Parse the XML
		$xml = new \SimpleXMLElement($this->temporaryDocumentMainPart);

		// Find the starting and ending tags
		$startNode = false;
		$endNode = false;
		foreach ($xml->xpath('//w:t') as $node) {
			if (strpos($node, '${' . $blockname . '}') !== false) {
				$startNode = $node;
				continue;
			}

			if (strpos($node, '${/' . $blockname . '}') !== false) {
				$endNode = $node;
				break;
			}
		}

		// Make sure we found the tags
		if ($startNode === false || $endNode === false) {
			//die('Tag "' . $blockname . '" not found in the document');
			return null;
		}

		// Find the parent <w:p> node for the start tag
		$node = $startNode;
		$startNode = null;
		while (is_null($startNode)) {
			$node = $node->xpath('..')[0];

			if ($node->getName() == 'p') {
				$startNode = $node;
			}
		}

		// Find the parent <w:p> node for the end tag
		$node = $endNode;
		$endNode = null;
		while (is_null($endNode)) {
			$node = $node->xpath('..')[0];

			if ($node->getName() == 'p') {
				$endNode = $node;
			}
		}

		$this->temporaryDocumentMainPart = $xml->asXml();

		// Find the xml in between the tags
		$xmlBlock = null;
		preg_match
		(
			'/' . preg_quote($startNode->asXml(), '/') . '(.*?)' . preg_quote($endNode->asXml(), '/') . '/is',
			$this->temporaryDocumentMainPart,
			$matches
		);

		if (isset($matches[1])) {
			$xmlBlock = $matches[1];

			$cloned = array();

			for ($i = 1; $i <= $clones; $i++) {
				$cloned[] = preg_replace('/\${(.*?)}/', '${$1_' . $i . '}', $xmlBlock);
			}

			if ($replace) {
				$this->temporaryDocumentMainPart = str_replace
				(
					$matches[0],
					implode('', $cloned),
					$this->temporaryDocumentMainPart
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
	 * @return void
	 */
	public function replaceBlock($blockname, $replacement) {
		// Parse the XML
		$xml = new \SimpleXMLElement($this->temporaryDocumentMainPart);

		// Find the starting and ending tags
		$startNode = false;
		$endNode = false;
		foreach ($xml->xpath('//w:t') as $node) {
			if (strpos($node, '${' . $blockname . '}') !== false) {
				$startNode = $node;
				continue;
			}

			if (strpos($node, '${/' . $blockname . '}') !== false) {
				$endNode = $node;
				break;
			}
		}

		// Make sure we found the tags
		if ($startNode === false || $endNode === false) {
			//die('Tag "' . $blockname . '" not found in the document');
			return null;
		}

		// Find the parent <w:p> node for the start tag
		$node = $startNode;
		$startNode = null;
		while (is_null($startNode)) {
			$node = $node->xpath('..')[0];

			if ($node->getName() == 'p') {
				$startNode = $node;
			}
		}

		// Find the parent <w:p> node for the end tag
		$node = $endNode;
		$endNode = null;
		while (is_null($endNode)) {
			$node = $node->xpath('..')[0];

			if ($node->getName() == 'p') {
				$endNode = $node;
			}
		}

		$this->temporaryDocumentMainPart = $xml->asXml();

		// Find the xml in between the tags
		$xmlBlock = null;
		preg_match
		(
			'/' . preg_quote($startNode->asXml(), '/') . '(.*?)' . preg_quote($endNode->asXml(), '/') . '/is',
			$this->temporaryDocumentMainPart,
			$matches
		);

		if (isset($matches[1])) {
			$this->temporaryDocumentMainPart = str_replace
			(
				$matches[0],
				$replacement,
				$this->temporaryDocumentMainPart
			);
		}
	}

	/**
	 * Delete a block of text.
	 *
	 * @param string $blockname
	 * @return void
	 */
	public function deleteBlock($blockname) {
		$this->replaceBlock($blockname, '');
	}

	/**
	 * Saves the result document.
	 *
	 * @return string
	 * @throws \PhpOffice\PhpWord\Exception\Exception
	 */
	public function save() {
		foreach ($this->temporaryDocumentHeaders as $index => $headerXML) {
			$this->zipClass->addFromString($this->getHeaderName($index), $this->temporaryDocumentHeaders[$index]);
		}

		$this->zipClass->addFromString('word/document.xml', str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $this->temporaryDocumentMainPart));

		foreach ($this->temporaryDocumentFooters as $index => $headerXML) {
			$this->zipClass->addFromString($this->getFooterName($index), $this->temporaryDocumentFooters[$index]);
		}

		// Close zip file
		if (false === $this->zipClass->close()) {
			throw new Exception('Could not close zip file.');
		}

		return $this->temporaryDocumentFilename;
	}

	/**
	 * Saves the result document to the user defined file.
	 *
	 * @since 0.8.0
	 *
	 * @param string $fileName
	 * @return void
	 */
	public function saveAs($fileName) {
		$tempFileName = $this->save();

		if (file_exists($fileName)) {
			unlink($fileName);
		}

		rename($tempFileName, $fileName);
	}

	/**
	 * Finds parts of broken macros and sticks them together.
	 * Macros, while being edited, could be implicitly broken by some of the word processors.
	 *
	 * @since 0.13.0
	 *
	 * @param string $documentPart The document part in XML representation.
	 *
	 * @return string
	 */
	protected function fixBrokenMacros($documentPart) {
		$fixedDocumentPart = $documentPart;

		$pattern = '|\$\{([^\}]+)\}|U';
		preg_match_all($pattern, $fixedDocumentPart, $matches);
		foreach ($matches[0] as $value) {
			$valueCleaned = preg_replace('/<[^>]+>/', '', $value);
			$valueCleaned = preg_replace('/<\/[^>]+>/', '', $valueCleaned);
			$fixedDocumentPart = str_replace($value, $valueCleaned, $fixedDocumentPart);
		}

		return $fixedDocumentPart;
	}

	/**
	 * Find and replace placeholders in the given XML section.
	 *
	 * @param string $documentPartXML
	 * @param string $search
	 * @param string $replace
	 * @param integer $limit
	 * @return string
	 */
	protected function setValueForPart($documentPartXML, $search, $replace, $limit) {
		if (substr($search, 0, 2) !== '${' && substr($search, -1) !== '}') {
			$search = '${' . $search . '}';
		}

		if (!WordString::isUTF8($replace)) {
			$replace = utf8_encode($replace);
		}

		$regExpDelim = '/';
		$escapedSearch = preg_quote($search, $regExpDelim);

		$found = false;
		if(strpos($replace, '<w:tbl>') === 0){
			$xml = new \SimpleXMLElement($documentPartXML);
			foreach ($xml->xpath("//w:p/*[contains(.,'{$search}')]/parent::*") as $node) {
				$escapedSearch = preg_quote($node->asXML(), $regExpDelim);
				$documentPartXML = preg_replace("{$regExpDelim}{$escapedSearch}{$regExpDelim}u", $replace, $documentPartXML, $limit);
				$found = true;
			}
			if($found){
				return $documentPartXML;
			}
		}

		return preg_replace("{$regExpDelim}{$escapedSearch}{$regExpDelim}u", $replace, $documentPartXML, $limit);
	}

	/**
	 * Find all variables in $documentPartXML.
	 *
	 * @param string $documentPartXML
	 * @return string[]
	 */
	protected function getVariablesForPart($documentPartXML) {
		preg_match_all('/\$\{(.*?)}/i', $documentPartXML, $matches);

		return $matches[1];
	}

	/**
	 * Get the name of the footer file for $index.
	 *
	 * @param integer $index
	 * @return string
	 */
	private function getFooterName($index) {
		return sprintf('word/footer%d.xml', $index);
	}

	/**
	 * Get the name of the header file for $index.
	 *
	 * @param integer $index
	 * @return string
	 */
	private function getHeaderName($index) {
		return sprintf('word/header%d.xml', $index);
	}

	/**
	 * Find the start position of the nearest table row before $offset.
	 *
	 * @param integer $offset
	 * @return integer
	 * @throws \PhpOffice\PhpWord\Exception\Exception
	 */
	private function findRowStart($offset) {
		$rowStart = strrpos($this->temporaryDocumentMainPart, '<w:tr ', ((strlen($this->temporaryDocumentMainPart) - $offset) * -1));

		if (!$rowStart) {
			$rowStart = strrpos($this->temporaryDocumentMainPart, '<w:tr>', ((strlen($this->temporaryDocumentMainPart) - $offset) * -1));
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
	 * @return integer
	 */
	private function findRowEnd($offset) {
		return strpos($this->temporaryDocumentMainPart, '</w:tr>', $offset) + 7;
	}

	/**
	 * Get a slice of a string.
	 *
	 * @param integer $startPosition
	 * @param integer $endPosition
	 * @return string
	 */
	private function getSlice($startPosition, $endPosition = 0) {
		if (!$endPosition) {
			$endPosition = strlen($this->temporaryDocumentMainPart);
		}

		return substr($this->temporaryDocumentMainPart, $startPosition, ($endPosition - $startPosition));
	}
}
