<?php
/**
 * This file is part of the PositibeLabs Projects.
 *
 * (c) Pedro Carlos Abreu <pcabreus@gmail.com>
 *
 * For the full copyright and license information, please view the LICENSE
 * file that was distributed with this source code.
 */

namespace Pcabreus\Utils\PhpWord;

use PhpOffice\PhpWord\Exception\Exception;
use PhpOffice\PhpWord\TemplateProcessor as OfficeTemplateProcessor;
use Positibe\Bundle\OrmMediaBundle\Entity\Media;

/**
 * Class TemplateProcessor
 * @package Pcabreus\PhpWord
 *
 * @author Pedro Carlos Abreu <pcabreus@gmail.com>
 */
class TemplateProcessor extends OfficeTemplateProcessor
{
    const SIGN_BLOCK = '<w:sym w:font="Wingdings" w:char="F06F"/>';
    const SIGN_EX = '<w:sym w:font="Wingdings 2" w:char="00D0"/>';

    protected $tempDocumentRels;
    protected $tempContentType;

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
        parent::__construct($documentTemplate);

        $this->tempDocumentRels = $this->zipClass->getFromName('word/_rels/document.xml.rels');
        $this->tempContentType = $this->zipClass->getFromName('[Content_Types].xml');
    }

    /**
     * @param $blockname
     * @param int $clones
     * @param bool|true $replace
     * @return null
     */
    public function cloneBlock($blockname, $clones = 1, $replace = true)
    {

        // Parse the XML
        $xml = new \SimpleXMLElement($this->tempDocumentMainPart);

        // Find the starting and ending tags
        $startNode = false;
        $endNode = false;
        foreach ($xml->xpath('//w:t') as $node) {
            if (strpos($node, '${'.$blockname.'}') !== false) {
                $startNode = $node;
                continue;
            }

            if (strpos($node, '${/'.$blockname.'}') !== false) {
                $endNode = $node;
                break;
            }
        }

        // Make sure we found the tags
        if ($startNode === false || $endNode === false) {
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

        /*
         * NOTE: Because SimpleXML reduces empty tags to "self-closing" tags.
         * We need to replace the original XML with the version of XML as
         * SimpleXML sees it. The following example should show the issue
         * we are facing.
         *
         * This is the XML that my document contained orginally.
         *
         * ```xml
         *  <w:p>
         *      <w:pPr>
         *          <w:pStyle w:val="TextBody"/>
         *          <w:rPr></w:rPr>
         *      </w:pPr>
         *      <w:r>
         *          <w:rPr></w:rPr>
         *          <w:t>${CLONEME}</w:t>
         *      </w:r>
         *  </w:p>
         * ```
         *
         * This is the XML that SimpleXML returns from asXml().
         *
         * ```xml
         *  <w:p>
         *      <w:pPr>
         *          <w:pStyle w:val="TextBody"/>
         *          <w:rPr/>
         *      </w:pPr>
         *      <w:r>
         *          <w:rPr/>
         *          <w:t>${CLONEME}</w:t>
         *      </w:r>
         *  </w:p>
         * ```
         */

        $this->tempDocumentMainPart = $xml->asXml();

        // Find the xml in between the tags
        $xmlBlock = null;
        preg_match
        (
            '/'.preg_quote($startNode->asXml(), '/').'(.*?)'.preg_quote($endNode->asXml(), '/').'/is',
            $this->tempDocumentMainPart,
            $matches
        );

        if (isset($matches[1])) {
            $xmlBlock = $matches[1];

            $cloned = array();

            for ($i = 1; $i <= $clones; $i++) {
                $cloned[] = preg_replace('/\${(.*?)}/', '${$1#'.$i.'}', $xmlBlock);
            }

            if ($replace) {
                $this->tempDocumentMainPart = str_replace
                (
                    $matches[0],
                    implode('', $cloned),
                    $this->tempDocumentMainPart
                );
            }
        }

        return $xmlBlock;
    }

    /**
     * @param $blockName
     * @param $filename
     * @param string $name
     * @param null $extension
     * @param null $type
     * @param null $width
     * @param null $height
     */
    public function setImage(
        $blockName,
        $filename,
        $name = 'Foto',
        $extension = null,
        $type = null,
        $width = null,
        $height = null
    ) {
        $image = $this->getProperties($filename, $name, $extension, $type, $width, $height);

        $this->setValue($blockName, $image['xmlImage']);

        $this->zipClass->addFile($filename, 'word/media/'.$image['media']);

        $this->tempDocumentRels = str_replace(
            '</Relationships>',
            $image['xmlRel'].'</Relationships>',
            $this->tempDocumentRels
        );

        if (!strpos($image['xmlType'], $this->tempContentType)) {
            $this->tempContentType = str_replace(
                '</Types>',
                $image['xmlType'].'</Types>',
                $this->tempContentType
            );

        }
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
        $this->tempDocumentMainPart = str_replace(
            ['<w:t><w:pict>', '</w:pict></w:t>'],
            ['<w:pict>', '</w:pict>'],
            $this->tempDocumentMainPart
        );

        $this->zipClass->addFromString('word/_rels/document.xml.rels', $this->tempDocumentRels);
        $this->zipClass->addFromString('[Content_Types].xml', $this->tempContentType);

        return parent::save();
    }


    private function getXmlId()
    {
        $xmlId = '_x0000_i'.rand(1000, 9999);
        if (strpos($xmlId, $this->tempDocumentMainPart)) {
            return self::getXmlId();
        }

        return $xmlId;
    }

    public function getRelsId()
    {
        $relsId = 'rId'.rand(50, 1000);
        if (strpos($relsId, $this->tempDocumentRels)) {
            return self::getRelsId();
        }

        return $relsId;
    }

    /**
     * @param $filename
     * @param string $name
     * @param null $extension
     * @param null $type
     * @param null $width
     * @param null $height
     * @return array
     */
    public function getProperties(
        $filename,
        $name = 'Image',
        $extension = null,
        $type = null,
        $width = null,
        $height = null
    ) {
        $xmlId = $this->getXmlId();
        $relsId = $this->getRelsId();

        if (!$extension) {
            $extension = strtolower(str_replace('.', '', strrchr($filename, '.')));
        }

        switch ($extension) {
            case 'jpg':
            case 'jpeg':
                $type = 'image/jpeg';
                $resource = @imagecreatefromjpeg($filename);
                break;
            case 'gif':
                $type = 'image/gir';
                $resource = @imagecreatefromgif($filename);
                break;
            case 'png':
                $type = 'image/png';
                $resource = @imagecreatefrompng($filename);
                break;
            default:
                $type = '';
                $resource = null;
                break;
        }
        $x = imagesx($resource);
        $y = imagesy($resource);
        if ($width && $height) {
            if ($x > $width && $x > $y) {
                $razon = $x / $width;
            } elseif ($y > $height && $y > $x) {
                $razon = $y / $height;
            } else {
                $razon = 1;
            }
            $width = $x / $razon;
            $height = $y / $razon;
        } else {
            $width = $x;
            $height = $y;
        }


        $media = 'image'.$relsId.'.'.$extension;

        $xmlImage = sprintf(
            '<w:pict><v:shape id="%s" type="#_x0000_t75" style="width:%spt;height:%spt"><v:imagedata r:id="%s" o:title="%s"/></v:shape></w:pict>',
            $xmlId,
            $width,
            $height,
            $relsId,
            $name
        );

        $xmlRel = sprintf(
            '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/%s"/>',
            $relsId,
            $media
        );

        return [
            'media' => $media,
            'width' => $width,
            'height' => $height,
            'xmlImage' => $xmlImage,
            'xmlRel' => $xmlRel,
            'xmlType' => sprintf('<Default Extension="%s" ContentType="%s"/>', $extension, $type),
        ];
    }

    /**
     * @param $blockName
     * @return null
     */
    public function getBlock($blockName)
    {
        preg_match(
            '/(<\?xml.*)(<w:p .*>\${'.$blockName.'}<\/w:.*?p>)(.*)(<w:p .*>\${\/'.$blockName.'}<\/w:.*?p>)/is',
            $this->tempDocumentMainPart,
            $matches
        );

        return isset($matches[3]) ? ['begin' => $matches[2], 'xmlBlock' => $matches[3], 'end' => $matches[4]] : null;
    }


    /**
     * @param $search
     * @param $clones Array of clones
     * @throws Exception
     */
    public function setValueBreakLine($search, $clones)
    {
        if (!is_array($clones)) {
            $clones = [$clones];
        }

        if ('${' !== substr($search, 0, 2) && '}' !== substr($search, -1)) {
            $search = '${'.$search.'}';
        }

        $tagPos = strpos($this->tempDocumentMainPart, $search);
        if (!$tagPos) {
            throw new Exception("Can not setValue row, template variable not found or variable contains markup.");
        }

        $rowStart = $this->findParagraphStart($tagPos);
        $rowEnd = $this->findParagraphEnd($tagPos);
        $xmlRow = $this->getSlice($rowStart, $rowEnd);

        $result = $this->getSlice(0, $rowStart);

        foreach ($clones as $clone) {
            $result .= str_replace($search, str_replace('&', '&amp;', $clone), $xmlRow);
        }

        $result .= $this->getSlice($rowEnd);

        $this->tempDocumentMainPart = $result;
    }

    /**
     * Find the start position of the nearest paragraph before $offset.
     *
     * @param integer $offset
     *
     * @return integer
     *
     * @throws \PhpOffice\PhpWord\Exception\Exception
     */
    protected function findParagraphStart($offset)
    {
        $rowStart = strrpos(
            $this->tempDocumentMainPart,
            '<w:p ',
            ((strlen($this->tempDocumentMainPart) - $offset) * -1)
        );

        if (!$rowStart) {
            $rowStart = strrpos(
                $this->tempDocumentMainPart,
                '<w:p >',
                ((strlen($this->tempDocumentMainPart) - $offset) * -1)
            );
        }
        if (!$rowStart) {
            throw new Exception('Can not find the start position of the paragraph to clone.');
        }

        return $rowStart;
    }

    /**
     * Find the end position of the nearest paragraph after $offset.
     *
     * @param integer $offset
     *
     * @return integer
     */
    protected function findParagraphEnd($offset)
    {
        return strpos($this->tempDocumentMainPart, '</w:p>', $offset) + 6;
    }
}