<?php
/**
 * Header file
*/
use PhpOffice\PhpPresentation\Autoloader;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\AbstractShape;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Group;
use PhpOffice\PhpPresentation\Shape\RichText;
use PhpOffice\PhpPresentation\Shape\RichText\BreakElement;
use PhpOffice\PhpPresentation\Shape\RichText\TextElement;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Bullet;
use PhpOffice\PhpPresentation\Style\Color;

error_reporting(E_ALL);
define('CLI', (PHP_SAPI == 'cli') ? true : false);
define('EOL', CLI ? PHP_EOL : '<br />');
define('SCRIPT_FILENAME', basename($_SERVER['SCRIPT_FILENAME'], '.php'));
define('IS_INDEX', SCRIPT_FILENAME == 'index');

require_once __DIR__ . '/../src/PhpPresentation/Autoloader.php';
Autoloader::register();

if (is_file(__DIR__. '/../../vendor/autoload.php')) {
    require_once __DIR__ . '/../../vendor/autoload.php';
} else {
    throw new Exception ('Can not find the vendor folder!');
}
// do some checks to make sure the outputs are set correctly.
if (is_dir(__DIR__.DIRECTORY_SEPARATOR.'results') === FALSE) {
    throw new Exception ('The results folder is not present!');
}
if (is_writable(__DIR__.DIRECTORY_SEPARATOR.'results'.DIRECTORY_SEPARATOR) === FALSE) {
    throw new Exception ('The results folder is not writable!');
}
if (is_writable(__DIR__.DIRECTORY_SEPARATOR) === FALSE) {
    throw new Exception ('The samples folder is not writable!');
}

// Set writers
$writers = array('PowerPoint2007' => 'pptx', 'ODPresentation' => 'odp');

// Return to the caller script when runs by CLI
if (CLI) {
    return;
}

// Set titles and names
$pageHeading = str_replace('_', ' ', SCRIPT_FILENAME);
$pageTitle = IS_INDEX ? 'Welcome to ' : "{$pageHeading} - ";
$pageTitle .= 'PHPPresentation';
$pageHeading = IS_INDEX ? '' : "<h1>{$pageHeading}</h1>";

$oShapeDrawing = new Drawing\File();
$oShapeDrawing->setName('PHPPresentation logo')
    ->setDescription('PHPPresentation logo')
    ->setPath('./resources/phppowerpoint_logo.gif')
    ->setHeight(36)
    ->setOffsetX(10)
    ->setOffsetY(10);
$oShapeDrawing->getShadow()->setVisible(true)
    ->setDirection(45)
    ->setDistance(10);
$oShapeDrawing->getHyperlink()->setUrl('https://github.com/PHPOffice/PHPPresentation/')->setTooltip('PHPPresentation');

// Create a shape (text)
$oShapeRichText = new RichText();
$oShapeRichText->setHeight(300)
    ->setWidth(600)
    ->setOffsetX(170)
    ->setOffsetY(180);
$oShapeRichText->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $oShapeRichText->createTextRun('Thank you for using PHPPresentation!');
$textRun->getFont()->setBold(true)
    ->setSize(60)
    ->setColor( new Color( 'FFE06B20' ) );



// Populate samples
$files = array();
if ($handle = opendir('.')) {
    while (false !== ($file = readdir($handle))) {
        if (preg_match('/^Sample_\d+_/', $file)) {
            $name = str_replace('_', ' ', preg_replace('/(Sample_|\.php)/', '', $file));
            $group = substr($name, 0, 1);
            if (!isset($files[$group])) {
                $files[$group] = '';
            }
            $files[$group] .= "<li><a href='{$file}'>{$name}</a></li>";
        }
    }
    closedir($handle);
}

/**
 * Write documents
 *
 * @param \PhpOffice\PhpPresentation\PhpPresentation $phpPresentation
 * @param string $filename
 * @param array $writers
 * @return string
 */
function write($phpPresentation, $filename, $writers)
{
    $result = '';
    
    // Write documents
    foreach ($writers as $writer => $extension) {
        $result .= date('H:i:s') . " Write to {$writer} format";
        if (!is_null($extension)) {
            $xmlWriter = IOFactory::createWriter($phpPresentation, $writer);
            $xmlWriter->save(__DIR__ . "/{$filename}.{$extension}");
            rename(__DIR__ . "/{$filename}.{$extension}", __DIR__ . "/results/{$filename}.{$extension}");
        } else {
            $result .= ' ... NOT DONE!';
        }
        $result .= EOL;
    }

    $result .= getEndingNotes($writers);

    return $result;
}

/**
 * Get ending notes
 *
 * @param array $writers
 * @return string
 */
function getEndingNotes($writers)
{
    $result = '';

    // Do not show execution time for index
    if (!IS_INDEX) {
        $result .= date('H:i:s') . " Done writing file(s)" . EOL;
        $result .= date('H:i:s') . " Peak memory usage: " . (memory_get_peak_usage(true) / 1024 / 1024) . " MB" . EOL;
    }

    // Return
    if (CLI) {
        $result .= 'The results are stored in the "results" subdirectory.' . EOL;
    } else {
        if (!IS_INDEX) {
            $types = array_values($writers);
            $result .= '<p>&nbsp;</p>';
            $result .= '<p>Results: ';
            foreach ($types as $type) {
                if (!is_null($type)) {
                    $resultFile = 'results/' . SCRIPT_FILENAME . '.' . $type;
                    if (file_exists($resultFile)) {
                        $result .= "<a href='{$resultFile}' class='btn btn-primary'>{$type}</a> ";
                    }
                }
            }
            $result .= '</p>';
        }
    }

    return $result;
}

/**
 * Creates a templated slide
 *
 * @param PHPPresentation $objPHPPresentation
 * @return \PhpOffice\PhpPresentation\Slide
 */
function createTemplatedSlide(PhpOffice\PhpPresentation\PhpPresentation $objPHPPresentation)
{
    // Create slide
    $slide = $objPHPPresentation->createSlide();
    
    // Add logo
    $shape = $slide->createDrawingShape();
    $shape->setName('PHPPresentation logo')
        ->setDescription('PHPPresentation logo')
        ->setPath('./resources/phppowerpoint_logo.gif')
        ->setHeight(36)
        ->setOffsetX(10)
        ->setOffsetY(10);
    $shape->getShadow()->setVisible(true)
        ->setDirection(45)
        ->setDistance(10);

    // Return slide
    return $slide;
}

class PhpPptTree {
    protected $oPhpPresentation;
    protected $htmlOutput;

    public function __construct(PhpPresentation $oPHPPpt)
    {
        $this->oPhpPresentation = $oPHPPpt;
    }

    public function display()
    {
        $this->append('<div class="container-fluid pptTree">');
        $this->append('<div class="row">');
        $this->append('<div class="collapse in col-md-6">');
        $this->append('<div class="tree">');
        $this->append('<ul>');
        $this->displayPhpPresentation($this->oPhpPresentation);
        $this->append('</ul>');
        $this->append('</div>');
        $this->append('</div>');
        $this->append('<div class="col-md-6">');
        $this->displayPhpPresentationInfo($this->oPhpPresentation);
        $this->append('</div>');
        $this->append('</div>');
        $this->append('</div>');

        return $this->htmlOutput;
    }

    protected function append($sHTML)
    {
        $this->htmlOutput .= $sHTML;
    }

    protected function displayPhpPresentation(PhpPresentation $oPHPPpt)
    {
        $this->append('<li><span><i class="fa fa-folder-open"></i> PhpPresentation</span>');
        $this->append('<ul>');
        $this->append('<li><span class="shape" id="divPhpPresentation"><i class="fa fa-info-circle"></i> Info "PhpPresentation"</span></li>');
        foreach ($oPHPPpt->getAllSlides() as $oSlide) {
            $this->append('<li><span><i class="fa fa-minus-square"></i> Slide</span>');
            $this->append('<ul>');
            $this->append('<li><span class="shape" id="div'.$oSlide->getHashCode().'"><i class="fa fa-info-circle"></i> Info "Slide"</span></li>');
            foreach ($oSlide->getShapeCollection() as $oShape) {
                if($oShape instanceof Group) {
                    $this->append('<li><span><i class="fa fa-minus-square"></i> Shape "Group"</span>');
                    $this->append('<ul>');
                    // $this->append('<li><span class="shape" id="div'.$oShape->getHashCode().'"><i class="fa fa-info-circle"></i> Info "Group"</span></li>');
                    foreach ($oShape->getShapeCollection() as $oShapeChild) {
                        $this->displayShape($oShapeChild);
                    }
                    $this->append('</ul>');
                    $this->append('</li>');
                } else {
                    $this->displayShape($oShape);
                }
            }
            $this->append('</ul>');
            $this->append('</li>');
        }
        $this->append('</ul>');
        $this->append('</li>');
    }

    protected function displayShape(AbstractShape $shape)
    {
        if($shape instanceof Drawing\Gd) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "Drawing\Gd"</span></li>');
        } elseif($shape instanceof Drawing\File) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "Drawing\File"</span></li>');
        } elseif($shape instanceof Drawing\Base64) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "Drawing\Base64"</span></li>');
        } elseif($shape instanceof Drawing\ZipFile) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "Drawing\Zip"</span></li>');
        } elseif($shape instanceof RichText) {
            $this->append('<li><span class="shape" id="div'.$shape->getHashCode().'">Shape "RichText"</span></li>');
        } else {
            var_dump($shape);
        }
    }

    protected function displayPhpPresentationInfo(PhpPresentation $oPHPPpt)
    {
        $this->append('<div class="infoBlk" id="divPhpPresentationInfo">');
        $this->append('<dl>');
        $this->append('<dt>Number of slides</dt><dd>'.$oPHPPpt->getSlideCount().'</dd>');
        $this->append('<dt>Document Layout Name</dt><dd>'.(empty($oPHPPpt->getLayout()->getDocumentLayout()) ? 'Custom' : $oPHPPpt->getLayout()->getDocumentLayout()).'</dd>');
        $this->append('<dt>Document Layout Height</dt><dd>'.$oPHPPpt->getLayout()->getCY(DocumentLayout::UNIT_MILLIMETER).' mm</dd>');
        $this->append('<dt>Document Layout Width</dt><dd>'.$oPHPPpt->getLayout()->getCX(DocumentLayout::UNIT_MILLIMETER).' mm</dd>');
        $this->append('<dt>Properties : Category</dt><dd>'.$oPHPPpt->getDocumentProperties()->getCategory().'</dd>');
        $this->append('<dt>Properties : Company</dt><dd>'.$oPHPPpt->getDocumentProperties()->getCompany().'</dd>');
        $this->append('<dt>Properties : Created</dt><dd>'.$oPHPPpt->getDocumentProperties()->getCreated().'</dd>');
        $this->append('<dt>Properties : Creator</dt><dd>'.$oPHPPpt->getDocumentProperties()->getCreator().'</dd>');
        $this->append('<dt>Properties : Description</dt><dd>'.$oPHPPpt->getDocumentProperties()->getDescription().'</dd>');
        $this->append('<dt>Properties : Keywords</dt><dd>'.$oPHPPpt->getDocumentProperties()->getKeywords().'</dd>');
        $this->append('<dt>Properties : Last Modified By</dt><dd>'.$oPHPPpt->getDocumentProperties()->getLastModifiedBy().'</dd>');
        $this->append('<dt>Properties : Modified</dt><dd>'.$oPHPPpt->getDocumentProperties()->getModified().'</dd>');
        $this->append('<dt>Properties : Subject</dt><dd>'.$oPHPPpt->getDocumentProperties()->getSubject().'</dd>');
        $this->append('<dt>Properties : Title</dt><dd>'.$oPHPPpt->getDocumentProperties()->getTitle().'</dd>');
        $this->append('</dl>');
        $this->append('</div>');

        foreach ($oPHPPpt->getAllSlides() as $oSlide) {
            $this->append('<div class="infoBlk" id="div'.$oSlide->getHashCode().'Info">');
            $this->append('<dl>');
            $this->append('<dt>HashCode</dt><dd>'.$oSlide->getHashCode().'</dd>');
            $this->append('<dt>Slide Layout</dt><dd>Layout::'.$this->getConstantName('\PhpOffice\PhpPresentation\Slide\Layout', $oSlide->getSlideLayout()).'</dd>');
            
            $this->append('<dt>Offset X</dt><dd>'.$oSlide->getOffsetX().'</dd>');
            $this->append('<dt>Offset Y</dt><dd>'.$oSlide->getOffsetY().'</dd>');
            $this->append('<dt>Extent X</dt><dd>'.$oSlide->getExtentX().'</dd>');
            $this->append('<dt>Extent Y</dt><dd>'.$oSlide->getExtentY().'</dd>');
            $oBkg = $oSlide->getBackground();
            if ($oBkg instanceof Slide\AbstractBackground) {
                if ($oBkg instanceof Slide\Background\Color) {
                    $this->append('<dt>Background Color</dt><dd>#'.$oBkg->getColor()->getRGB().'</dd>');
                }
                if ($oBkg instanceof Slide\Background\Image) {
                    $sBkgImgContents = file_get_contents($oBkg->getPath());
                    $this->append('<dt>Background Image</dt><dd><img src="data:image/png;base64,'.base64_encode($sBkgImgContents).'"></dd>');
                }
            }
            $oNote = $oSlide->getNote();
            if ($oNote->getShapeCollection()->count() > 0) {
                $this->append('<dt>Notes</dt>');
                foreach ($oNote->getShapeCollection() as $oShape) {
                    if ($oShape instanceof RichText) {
                        $this->append('<dd>' . $oShape->getPlainText() . '</dd>');
                    }
                }
            }

            $this->append('</dl>');
            $this->append('</div>');

            foreach ($oSlide->getShapeCollection() as $oShape) {
                if($oShape instanceof Group) {
                    foreach ($oShape->getShapeCollection() as $oShapeChild) {
                        $this->displayShapeInfo($oShapeChild);
                    }
                } else {
                    $this->displayShapeInfo($oShape);
                }
            }
        }
    }

    protected function displayShapeInfo(AbstractShape $oShape)
    {
        $this->append('<div class="infoBlk" id="div'.$oShape->getHashCode().'Info">');
        $this->append('<dl>');
        $this->append('<dt>HashCode</dt><dd>'.$oShape->getHashCode().'</dd>');
        $this->append('<dt>Offset X</dt><dd>'.$oShape->getOffsetX().'</dd>');
        $this->append('<dt>Offset Y</dt><dd>'.$oShape->getOffsetY().'</dd>');
        $this->append('<dt>Height</dt><dd>'.$oShape->getHeight().'</dd>');
        $this->append('<dt>Width</dt><dd>'.$oShape->getWidth().'</dd>');
        $this->append('<dt>Rotation</dt><dd>'.$oShape->getRotation().'°</dd>');
        $this->append('<dt>Hyperlink</dt><dd>'.ucfirst(var_export($oShape->hasHyperlink(), true)).'</dd>');
        $this->append('<dt>Fill</dt>');
        if (is_null($oShape->getFill())) {
            $this->append('<dd>None</dd>');
        } else {
            switch($oShape->getFill()->getFillType()) {
                case \PhpOffice\PhpPresentation\Style\Fill::FILL_NONE:
                    $this->append('<dd>None</dd>');
                    break;
                case \PhpOffice\PhpPresentation\Style\Fill::FILL_SOLID:
                    $this->append('<dd>Solid (');
                    $this->append('Color : #'.$oShape->getFill()->getStartColor()->getRGB());
                    $this->append(' - Alpha : '.$oShape->getFill()->getStartColor()->getAlpha().'%');
                    $this->append(')</dd>');
                    break;
            }
        }
        $this->append('<dt>Border</dt><dd>@Todo</dd>');
        $this->append('<dt>IsPlaceholder</dt><dd>' . ($oShape->isPlaceholder() ? 'true' : 'false') . '</dd>');
        if($oShape instanceof Drawing\Gd) {
            $this->append('<dt>Name</dt><dd>'.$oShape->getName().'</dd>');
            $this->append('<dt>Description</dt><dd>'.$oShape->getDescription().'</dd>');
            ob_start();
            call_user_func($oShape->getRenderingFunction(), $oShape->getImageResource());
            $sShapeImgContents = ob_get_contents();
            ob_end_clean();
            $this->append('<dt>Mime-Type</dt><dd>'.$oShape->getMimeType().'</dd>');
            $this->append('<dt>Image</dt><dd><img src="data:'.$oShape->getMimeType().';base64,'.base64_encode($sShapeImgContents).'"></dd>');
        } elseif($oShape instanceof Drawing\AbstractDrawingAdapter) {
            $this->append('<dt>Name</dt><dd>'.$oShape->getName().'</dd>');
            $this->append('<dt>Description</dt><dd>'.$oShape->getDescription().'</dd>');
        } elseif($oShape instanceof RichText) {
            $this->append('<dt># of paragraphs</dt><dd>'.count($oShape->getParagraphs()).'</dd>');
            $this->append('<dt>Inset (T / R / B / L)</dt><dd>'.$oShape->getInsetTop().'px / '.$oShape->getInsetRight().'px / '.$oShape->getInsetBottom().'px / '.$oShape->getInsetLeft().'px</dd>');
            $this->append('<dt>Text</dt>');
            $this->append('<dd>');
            foreach ($oShape->getParagraphs() as $oParagraph) {
                $this->append('Paragraph<dl>');
                $this->append('<dt>Alignment Horizontal</dt><dd> Alignment::'.$this->getConstantName('\PhpOffice\PhpPresentation\Style\Alignment', $oParagraph->getAlignment()->getHorizontal()).'</dd>');
                $this->append('<dt>Alignment Vertical</dt><dd> Alignment::'.$this->getConstantName('\PhpOffice\PhpPresentation\Style\Alignment', $oParagraph->getAlignment()->getVertical()).'</dd>');
                $this->append('<dt>Alignment Margin (L / R)</dt><dd>'.$oParagraph->getAlignment()->getMarginLeft().' px / '.$oParagraph->getAlignment()->getMarginRight().'px</dd>');
                $this->append('<dt>Alignment Indent</dt><dd>'.$oParagraph->getAlignment()->getIndent().' px</dd>');
                $this->append('<dt>Alignment Level</dt><dd>'.$oParagraph->getAlignment()->getLevel().'</dd>');
                $this->append('<dt>Bullet Style</dt><dd> Bullet::'.$this->getConstantName('\PhpOffice\PhpPresentation\Style\Bullet', $oParagraph->getBulletStyle()->getBulletType()).'</dd>');
                if ($oParagraph->getBulletStyle()->getBulletType() != Bullet::TYPE_NONE) {
                    $this->append('<dt>Bullet Font</dt><dd>' . $oParagraph->getBulletStyle()->getBulletFont() . '</dd>');
                    $this->append('<dt>Bullet Color</dt><dd>' . $oParagraph->getBulletStyle()->getBulletColor()->getARGB() . '</dd>');
                }
                if ($oParagraph->getBulletStyle()->getBulletType() == Bullet::TYPE_BULLET) {
                    $this->append('<dt>Bullet Char</dt><dd>'.$oParagraph->getBulletStyle()->getBulletChar().'</dd>');
                }
                if ($oParagraph->getBulletStyle()->getBulletType() == Bullet::TYPE_NUMERIC) {
                    $this->append('<dt>Bullet Start At</dt><dd>'.$oParagraph->getBulletStyle()->getBulletNumericStartAt().'</dd>');
                    $this->append('<dt>Bullet Style</dt><dd>'.$oParagraph->getBulletStyle()->getBulletNumericStyle().'</dd>');
                }
                $this->append('<dt>Line Spacing</dt><dd>'.$oParagraph->getLineSpacing().'</dd>');
                $this->append('<dt>RichText</dt><dd><dl>');
                foreach ($oParagraph->getRichTextElements() as $oRichText) {
                    if($oRichText instanceof BreakElement) {
                        $this->append('<dt><i>Break</i></dt>');
                    } else {
                        if ($oRichText instanceof TextElement) {
                           $this->append('<dt><i>TextElement</i></dt>');
                        } else {
                           $this->append('<dt><i>Run</i></dt>');
                        }
                        $this->append('<dd>'.$oRichText->getText());
                        $this->append('<dl>');
                        $this->append('<dt>Font Name</dt><dd>'.$oRichText->getFont()->getName().'</dd>');
                        $this->append('<dt>Font Size</dt><dd>'.$oRichText->getFont()->getSize().'</dd>');
                        $this->append('<dt>Font Color</dt><dd>#'.$oRichText->getFont()->getColor()->getARGB().'</dd>');
                        $this->append('<dt>Font Transform</dt><dd>');
                            $this->append('<abbr title="Bold">Bold</abbr> : '.($oRichText->getFont()->isBold() ? 'Y' : 'N').' - ');
                            $this->append('<abbr title="Italic">Italic</abbr> : '.($oRichText->getFont()->isItalic() ? 'Y' : 'N').' - ');
                            $this->append('<abbr title="Underline">Underline</abbr> : Underline::'.$this->getConstantName('\PhpOffice\PhpPresentation\Style\Font', $oRichText->getFont()->getUnderline()).' - ');
                            $this->append('<abbr title="Strikethrough">Strikethrough</abbr> : '.($oRichText->getFont()->isStrikethrough() ? 'Y' : 'N').' - ');
                            $this->append('<abbr title="SubScript">SubScript</abbr> : '.($oRichText->getFont()->isSubScript() ? 'Y' : 'N').' - ');
                            $this->append('<abbr title="SuperScript">SuperScript</abbr> : '.($oRichText->getFont()->isSuperScript() ? 'Y' : 'N'));
                        $this->append('</dd>');
                        if ($oRichText instanceof TextElement) {
                            if ($oRichText->hasHyperlink()) {
                                $this->append('<dt>Hyperlink URL</dt><dd>'.$oRichText->getHyperlink()->getUrl().'</dd>');
                                $this->append('<dt>Hyperlink Tooltip</dt><dd>'.$oRichText->getHyperlink()->getTooltip().'</dd>');
                            }
                        }
                        $this->append('</dl>');
                        $this->append('</dd>');
                    }
                }
                $this->append('</dl></dd></dl>');
            }
            $this->append('</dd>');
        } else {
            // Add another shape
        }
        $this->append('</dl>');
        $this->append('</div>');
    }
    
    protected function getConstantName($class, $search, $startWith = '') {
        $fooClass = new ReflectionClass($class);
        $constants = $fooClass->getConstants();
        $constName = null;
        foreach ($constants as $key => $value ) {
            if ($value == $search) {
                if (empty($startWith) || (!empty($startWith) && strpos($key, $startWith) === 0)) {
                    $constName = $key;
                }
                break;
            }
        }
        return $constName;
    }
}
?>

<!DOCTYPE html>  
 <html>  
       <head>
	  <meta charset="utf-8">
	  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
	  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
	  <meta name="viewport" content="width=device-width, initial-scale=1">
      <link href="../dist/css/lightgallery.css" rel="stylesheet">
	  <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.2/jquery.min.js"></script> 
	   <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.1.0/jquery.min.js"></script>  
	   <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" />  
	   <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>  
		   
		<style type="text/css">

            body {
              font-family: Arial, Helvetica, sans-serif;
            }
            
            .demo-gallery > ul {
              margin-bottom: 0;
            }
            .demo-gallery > ul > li {
                float: left;
                margin-bottom: 15px;
                margin-right: 20px;
                width: 200px;
            }
            .demo-gallery > ul > li a {
              border: 3px solid #FFF;
              border-radius: 3px;
              display: block;
              overflow: hidden;
              position: relative;
              float: left;
            }
            .demo-gallery > ul > li a > img {
              -webkit-transition: -webkit-transform 0.15s ease 0s;
              -moz-transition: -moz-transform 0.15s ease 0s;
              -o-transition: -o-transform 0.15s ease 0s;
              transition: transform 0.15s ease 0s;
              -webkit-transform: scale3d(1, 1, 1);
              transform: scale3d(1, 1, 1);
              height: 100%;
              width: 100%;
            }
            .demo-gallery > ul > li a:hover > img {
              -webkit-transform: scale3d(1.1, 1.1, 1.1);
              transform: scale3d(1.1, 1.1, 1.1);
            }
            .demo-gallery > ul > li a:hover .demo-gallery-poster > img {
              opacity: 1;
            }
            .demo-gallery > ul > li a .demo-gallery-poster {
              background-color: rgba(0, 0, 0, 0.1);
              bottom: 0;
              left: 0;
              position: absolute;
              right: 0;
              top: 0;
              -webkit-transition: background-color 0.15s ease 0s;
              -o-transition: background-color 0.15s ease 0s;
              transition: background-color 0.15s ease 0s;
            }
            .demo-gallery > ul > li a .demo-gallery-poster > img {
              left: 50%;
              margin-left: -10px;
              margin-top: -10px;
              opacity: 0;
              position: absolute;
              top: 50%;
              -webkit-transition: opacity 0.3s ease 0s;
              -o-transition: opacity 0.3s ease 0s;
              transition: opacity 0.3s ease 0s;
            }
            .demo-gallery > ul > li a:hover .demo-gallery-poster {
              background-color: rgba(0, 0, 0, 0.5);
            }
            .demo-gallery .justified-gallery > a > img {
              -webkit-transition: -webkit-transform 0.15s ease 0s;
              -moz-transition: -moz-transform 0.15s ease 0s;
              -o-transition: -o-transform 0.15s ease 0s;
              transition: transform 0.15s ease 0s;
              -webkit-transform: scale3d(1, 1, 1);
              transform: scale3d(1, 1, 1);
              height: 100%;
              width: 100%;
            }
            .demo-gallery .justified-gallery > a:hover > img {
              -webkit-transform: scale3d(1.1, 1.1, 1.1);
              transform: scale3d(1.1, 1.1, 1.1);
            }
            .demo-gallery .justified-gallery > a:hover .demo-gallery-poster > img {
              opacity: 1;
            }
            .demo-gallery .justified-gallery > a .demo-gallery-poster {
              background-color: rgba(0, 0, 0, 0.1);
              bottom: 0;
              left: 0;
              position: absolute;
              right: 0;
              top: 0;
              -webkit-transition: background-color 0.15s ease 0s;
              -o-transition: background-color 0.15s ease 0s;
              transition: background-color 0.15s ease 0s;
            }
            .demo-gallery .justified-gallery > a .demo-gallery-poster > img {
              left: 50%;
              margin-left: -10px;
              margin-top: -10px;
              opacity: 0;
              position: absolute;
              top: 50%;
              -webkit-transition: opacity 0.3s ease 0s;
              -o-transition: opacity 0.3s ease 0s;
              transition: opacity 0.3s ease 0s;
            }
            .demo-gallery .justified-gallery > a:hover .demo-gallery-poster {
              background-color: rgba(0, 0, 0, 0.5);
            }
            .demo-gallery .video .demo-gallery-poster img {
              height: 48px;
              margin-left: -24px;
              margin-top: -24px;
              opacity: 0.8;
              width: 48px;
            }
            .demo-gallery.dark > ul > li a {
              border: 3px solid #04070a;
            }
            .home .demo-gallery {
              padding-bottom: 80px;
            }

           
            ul {
              list-style-type: none;
              margin: 0;
              padding: 0;
              overflow: hidden;
              width: 100%;
              background-color: white;
            }

            ul {
              float: left;
            }

            li a, .dropbtn {
              display: inline-block;
              color: white;
              text-align: center;
              padding: 16px 16px;
              text-decoration: none;
            }

            li a:hover, .dropdown:hover .dropbtn {
              background-color: #f2f2f2;
            }

            li.dropdown {
              display: inline-block;
            }

            .dropdown-content {
              display: none;
              position: absolute;
              background-color: #f2f2f2;
              min-width: 160px;
              box-shadow: 0px 8px 16px 0px rgba(0,0,0,0.2);
              z-index: 1;
            }

            .dropdown-content a {
              color: #595959;
              padding: 14px 16px;
              text-decoration: none;
              display: block;
              text-align: left;
            }

            .dropdown-content a:hover {background-color: #e6e6e6
            ;}

            .dropdown:hover .dropdown-content {
              display: block;
            }

            * {
              box-sizing: border-box;
            }

           table {
                border-collapse: collapse;
                border: 2px black solid;
                font: 16px sans-serif;
                text-align: center;
            }

            td {
                border: 1px black solid;
                padding: 5px;
                padding-left: 10px;
                padding-right: 10px;

            }    
        </style>
      </head>  
      <body class="home">		
		<div>
        <ul>
        <li><a style="float:left; color:#8c8c8c; font-size: 16px;">IMAGE PROCESSING BASED SMART SURVEILLANCE SYSTEM</a></li>
		<li><a href="../../../ln.html" style="float:right; color:#8c8c8c; font-size: 16px;">Logout</a></li>
		<li><a href="../../lightGallery-master/demo/HOME1.html" style="float:right; color:#8c8c8c; font-size: 16px;">Home</a></li>
		<li><a href="../../lightGallery-master/demo/excel.php" style="float:right; color:#8c8c8c; font-size: 16px;">Excel</a></li>
		
        
      </ul>
      </div>
	  <center><div>
      <img src="e.png" style="width: 200px; padding-top: 20px;" />
    <center><p style="font-size: 36px; color: #1a1a1a; ">REPORT OF SYSTEM</p></center>
	</div></center>
	</body>
	<footer>
     
           <div class="container" style="width:98%;">   
			
               
           </div>  
		<script src="https://cdn.jsdelivr.net/picturefill/2.3.1/picturefill.min.js"></script>
        <script src="../dist/js/lightgallery-all.min.js"></script>
        <script src="../lib/jquery.mousewheel.min.js"></script>
      </footer>  
 </html>  
 



