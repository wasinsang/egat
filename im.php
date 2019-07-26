<?php

include_once 'Sample_Header.php';


use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;

// Create new PHPPresentation object
$objPHPPresentation = new PhpPresentation();

// Remove first slide
$objPHPPresentation->removeSlideByIndex(0);

// Create templated slide 1

$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

// Create a shape (text)

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100);
$shape->setWidth(600);
$shape->setOffsetX(200);
$shape->setOffsetY(200);

$textRun = $shape->createTextRun('MONTHLY  REPORT ');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(48);

$shape->createBreak();

// Create a shape (text)

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100);
$shape->setWidth(800);
$shape->setOffsetX(130);
$shape->setOffsetY(320);

$textRun = $shape->createTextRun('Image Processing Based Smart');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(40);

$shape->createBreak();

// Create a shape (text)

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100);
$shape->setWidth(600);
$shape->setOffsetX(235);
$shape->setOffsetY(400);

$textRun = $shape->createTextRun('Surveillance System');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(40);



$folderPath = 'img';
foreach(glob($folderPath.'/*.jpg') as $file) {
	
$img = "./$file";


// Create templated slide 1

$currentSlide = createTemplatedSlide($objPHPPresentation); // local function

// Generate an image

$gdImage = @imagecreatetruecolor(450, 40) or die('Cannot Initialize new GD image stream');

// Add a generated drawing to the slide

$shape = new Drawing\Gd();
$shape->setName('Sample image')
      ->setDescription('Sample image')
      ->setImageResource($gdImage)
      ->setMimeType(Drawing\Gd::MIMETYPE_DEFAULT)
      ->setHeight(60)
      ->setOffsetX(5)
      ->setOffsetY(5);
	  
$currentSlide->addShape($shape);


// Create a shape (text)

$shape = $currentSlide->createRichTextShape();
$shape->setHeight(100);
$shape->setWidth(600);
$shape->setOffsetX(25);
$shape->setOffsetY(15);

$textRun = $shape->createTextRun('Image Processing Based Smart Surveillance System Report');
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(18);
$textRun->getFont()->setColor(new Color('f2f2f2'));

$shape->createBreak();


$shape = $currentSlide->createRichTextShape();
$shape->setHeight(200);
$shape->setWidth(600);
$shape->setOffsetX(200);
$shape->setOffsetY(90);

$name = substr($file,4,15);
$dated = substr($file,24,2);
$mounth = substr($file,22,2);
$year = substr($file,20,2);
$hour = substr($file,27,2);
$min = substr($file,29,2);

$textRun = $shape->createTextRun("Image Name: $name");
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(28);
//$textRun->getFont()->setColor(new Color('00ff0080'));
$shape->createBreak();


$shape = $currentSlide->createRichTextShape();
$shape->setHeight(200);
$shape->setWidth(600);
$shape->setOffsetX(270);
$shape->setOffsetY(150);
$textRun = $shape->createTextRun("Date: $dated/$mounth/$year   Time  : $hour:$min");
$textRun->getFont()->setBold(true);
$textRun->getFont()->setSize(20);
//$textRun->getFont()->setColor(new Color('00ff0080'));
$shape->createBreak();


// Add a file drawing (GIF) to the slide
$shape = new Drawing\File();
$shape->setPath($img)
    ->setHeight(460)
    ->setOffsetX(150)
    ->setOffsetY(215);	
$currentSlide->addShape($shape);


}
?>
<center><?php

echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Sample_Footer.php';
}

?></center>