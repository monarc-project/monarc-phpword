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

namespace PhpOffice\PhpWord\Tests\Element;

use PhpOffice\PhpWord\Element\ListItemRun;

/**
 * Test class for PhpOffice\PhpWord\Element\ListItemRun
 *
 * @runTestsInSeparateProcesses
 */
class ListItemRunTest extends \PHPUnit_Framework_TestCase
{
    /**
     * New instance
     */
    public function testConstructNull()
    {
        $oListItemRun = new ListItemRun();

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\ListItemRun', $oListItemRun);
        $this->assertCount(0, $oListItemRun->getElements());
        $this->assertNull($oListItemRun->getParagraphStyle());
    }

    /**
     * New instance with string
     */
    public function testConstructString()
    {
        $oListItemRun = new ListItemRun(0, null, 'pStyle');

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\ListItemRun', $oListItemRun);
        $this->assertCount(0, $oListItemRun->getElements());
        $this->assertEquals('pStyle', $oListItemRun->getParagraphStyle());
    }

    /**
     * New instance with string
     */
    public function testConstructListString()
    {
        $oListItemRun = new ListItemRun(0, 'numberingStyle');

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\ListItemRun', $oListItemRun);
        $this->assertCount(0, $oListItemRun->getElements());
    }

    /**
     * New instance with array
     */
    public function testConstructArray()
    {
        $oListItemRun = new ListItemRun(0, null, array('spacing' => 100));

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\ListItemRun', $oListItemRun);
        $this->assertCount(0, $oListItemRun->getElements());
        $this->assertInstanceOf('PhpOffice\\PhpWord\\Style\\Paragraph', $oListItemRun->getParagraphStyle());
    }

    /**
     * Get style
     */
    public function testStyle()
    {
        $oListItemRun = new ListItemRun(1, array('listType' => \PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER));

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Style\\ListItem', $oListItemRun->getStyle());
        $this->assertEquals(\PhpOffice\PhpWord\Style\ListItem::TYPE_NUMBER, $oListItemRun->getStyle()->getListType());
    }

    /**
     * getDepth
     */
    public function testDepth()
    {
        $iVal = rand(1, 1000);
        $oListItemRun = new ListItemRun($iVal);

        $this->assertEquals($iVal, $oListItemRun->getDepth());
    }

    /**
     * Add text
     */
    public function testAddText()
    {
        $oListItemRun = new ListItemRun();
        $element = $oListItemRun->addText(htmlspecialchars('text', ENT_COMPAT, 'UTF-8'));

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\Text', $element);
        $this->assertCount(1, $oListItemRun->getElements());
        $this->assertEquals(htmlspecialchars('text', ENT_COMPAT, 'UTF-8'), $element->getText());
    }

    /**
     * Add text non-UTF8
     */
    public function testAddTextNotUTF8()
    {
        $oListItemRun = new ListItemRun();
        $element = $oListItemRun->addText(utf8_decode(htmlspecialchars('ééé', ENT_COMPAT, 'UTF-8')));

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\Text', $element);
        $this->assertCount(1, $oListItemRun->getElements());
        $this->assertEquals(htmlspecialchars('ééé', ENT_COMPAT, 'UTF-8'), $element->getText());
    }

    /**
     * Add link
     */
    public function testAddLink()
    {
        $oListItemRun = new ListItemRun();
        $element = $oListItemRun->addLink('https://github.com/PHPOffice/PHPWord');

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\Link', $element);
        $this->assertCount(1, $oListItemRun->getElements());
        $this->assertEquals('https://github.com/PHPOffice/PHPWord', $element->getSource());
    }

    /**
     * Add link with name
     */
    public function testAddLinkWithName()
    {
        $oListItemRun = new ListItemRun();
        $element = $oListItemRun->addLink('https://github.com/PHPOffice/PHPWord', htmlspecialchars('PHPWord on GitHub', ENT_COMPAT, 'UTF-8'));

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\Link', $element);
        $this->assertCount(1, $oListItemRun->getElements());
        $this->assertEquals('https://github.com/PHPOffice/PHPWord', $element->getSource());
        $this->assertEquals(htmlspecialchars('PHPWord on GitHub', ENT_COMPAT, 'UTF-8'), $element->getText());
    }

    /**
     * Add text break
     */
    public function testAddTextBreak()
    {
        $oListItemRun = new ListItemRun();
        $oListItemRun->addTextBreak(2);

        $this->assertCount(2, $oListItemRun->getElements());
    }

    /**
     * Add image
     */
    public function testAddImage()
    {
        $src = __DIR__ . '/../_files/images/earth.jpg';

        $oListItemRun = new ListItemRun();
        $element = $oListItemRun->addImage($src);

        $this->assertInstanceOf('PhpOffice\\PhpWord\\Element\\Image', $element);
        $this->assertCount(1, $oListItemRun->getElements());
    }
}
