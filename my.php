<?php
require_once 'PHPPresentation/src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
require_once 'Common/src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();
require_once 'vendor/autoload.php';
include_once 'Header.php';
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Shape\Drawing;
use PhpOffice\PhpPresentation\Shape\Media;
$servername = "localhost";
$username = "root";
$password = "1234";
$dbname = "egat";
$colorBlack = new Color( 'FF0868AC' );
$ID = array();
$NAME = array();
$DATE = array();
$TIME = array();
$DETECTION = array();
$URL = array();
$currentSlide = array();
$shape = array();
$I = 0;
$A = 0;
$ALL = 0;
// Create new PHPPresentation object
		$objPHPPresentation = new PhpPresentation();
// Set properties
		$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPresentation Team')
                                  ->setTitle('Sample 06 Title')
                                  ->setSubject('Sample 06 Subject')
                                  ->setDescription('Sample 06 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
		$objPHPPresentation->removeSlideByIndex(0);
		
// Create connection
$conn = new mysqli($servername, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 
mysqli_set_charset($conn,"utf8");
//echo "Connected successfully";
echo "<br>";
$sql = "SELECT id, name, date,time, detection, url FROM egatreportnewmaster";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
    // output data of each row
    while($row = $result->fetch_assoc()) {
        //echo "id: " . $row["id"].  $row["name"].  $row["date"].  $row["time"].  $row["detection"]. "<br>";
		$ID[$I] = $row["id"];
		$NAME[$I] = $row["name"];
		$DATE[$I] = $row["date"];
		$TIME[$I] = $row["time"];
		$DETECTION[$I] = $row["detection"];
		$url[$I] = $row["url"];
		//Get the file
		$content = file_get_contents($url[$I]);
		//Store in the filesystem.
		$fp = fopen("image/image$I.jpg", "w");
		fwrite($fp, $content);
		fclose($fp);
		$I ++;
    }
} else {
    echo "0 results";
}
$conn->close();
$ALL = count($ID);
//echo $ALL;
$cam = 0;
for ($C = 0; $C <= $ALL-1; $C++) 
{
	for ($D = $C+1; $D <= $ALL-1; $D++) 
	{
		if($ID[$C] == $ID[$D])
		{
			$cam ++;
		}
	}
}
// Create new PHPPresentation object
$objPHPPresentation = new PhpPresentation();

// Set properties
$objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
                                  ->setLastModifiedBy('PHPPresentation Team')
                                  ->setTitle('Sample 06 Title')
                                  ->setSubject('Sample 06 Subject')
                                  ->setDescription('Sample 06 Description')
                                  ->setKeywords('office 2007 openxml libreoffice odt php')
                                  ->setCategory('Sample Category');

// Remove first slide
$objPHPPresentation->removeSlideByIndex(0);
// Create slide
$currentSlide = createTemplatedSlide($objPHPPresentation);
// Add a file drawing (GIF) to the slide
$shape = new Drawing\File();
$shape->setName('PHPPresentation logo')
    ->setDescription('PHPPresentation logo')
    ->setPath('image/back.jpg')
    ->setHeight(750)
    ->setOffsetX(0)
    ->setOffsetY(0);
$currentSlide->addShape($shape);
// Create a shape (text)
$shape = $currentSlide->createRichTextShape()
    ->setHeight(300)
    ->setWidth(650)
    ->setOffsetX(200)
    ->setOffsetY(250);
$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
$textRun = $shape->createTextRun('รายงานผลการดำเนินงาน สิ่งปลูกสร้างรุกล้ำแนวเขตเดินสายส่งไฟฟ้า');
$textRun->getFont()->setBold(true)
    ->setSize(50)
    ->setColor(new Color('FFFFFFFF'));
	
// Create slide
$currentSlide = createTemplatedSlide($objPHPPresentation);

// Create a shape (text)
	$shape = $currentSlide->createRichTextShape()
		->setHeight(50)
		->setWidth(800)
		->setOffsetX(150)
		->setOffsetY(80);
	$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
	$textRun = $shape->createTextRun('รายละเอียดเหตุการณ์ที่ป้องกันได้ประจำเดือน 2562');
	$textRun->getFont()->setBold(true)
		->setSize(24)
		->setColor($colorBlack);
// Create a shape (table)	
	$shape = $currentSlide->createTableShape(7);
	$shape->setHeight(400);
	$shape->setWidth(800);
	$shape->setOffsetX(80);
	$shape->setOffsetY(200);
// Add row
	$row = $shape->createRow();
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffff00'))
				   ->setEndColor(new Color('ffff00'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('กสx-ส')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	$oCell = $row->nextCell();
	$oCell->createTextRun('จำนวนเหตุการณ์')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	$oCell = $row->nextCell();
	$oCell->createTextRun('Truck')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	$oCell = $row->nextCell();
	$oCell->createTextRun('Backhoe')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	$oCell = $row->nextCell();
	$oCell->createTextRun('Crane')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	$oCell = $row->nextCell();
	$oCell->createTextRun('อื่นๆ')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	$oCell = $row->nextCell();
	$oCell->createTextRun('หมายเหตุ')->getFont()->setBold(true)->setSize(20)->setColor($colorBlack);;
	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffe6cc'))
				   ->setEndColor(new Color('ffe6cc'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('กสล-ส.')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment();
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffffff'))
				   ->setEndColor(new Color('ffffff'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('กสก-ส.')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment();
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffe6cc'))
				   ->setEndColor(new Color('ffe6cc'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('กสอ-ส.')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment();
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffffff'))
				   ->setEndColor(new Color('ffffff'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('กสต-ส.')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment();
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffe6cc'))
				   ->setEndColor(new Color('ffe6cc'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('กสน-ส.')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment();
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffffff'))
				   ->setEndColor(new Color('ffffff'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('รวม')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment();
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ALL)->getFont()->setSize(14);

for ($x = 0; $x <= $ALL-1; $x++) 
{
	// Create slide
	$currentSlide = createTemplatedSlide($objPHPPresentation);

	// Create a shape (table)	
	$shape = $currentSlide->createTableShape(2);
	$shape->setHeight(200);
	$shape->setWidth(400);
	$shape->setOffsetX(30);
	$shape->setOffsetY(200);

	// Add row
	$row = $shape->createRow();
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffff00'))
				   ->setEndColor(new Color('ffff00'));
	$cell = $row->nextCell();
	$cell->setColSpan(2);
	$cell->createTextRun('รายละเอียดเหตุการณ์')->getFont()->setBold(true)->setSize(18);
	$cell->getActiveParagraph()->getAlignment()
		->setMarginLeft(70);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffe6cc'))
				   ->setEndColor(new Color('ffe6cc'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('ID')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(50);
	$oCell = $row->nextCell();
	$oCell->createTextRun($ID[$A])->getFont()->setSize(14);
	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffffff'))
				   ->setEndColor(new Color('ffffff'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('NAME')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(40);
	$oCell = $row->nextCell();
	$oCell->createTextRun($NAME[$A])->getFont()->setSize(14);
	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffe6cc'))
				   ->setEndColor(new Color('ffe6cc'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('DATE AND TIME')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10);
	$oCell = $row->nextCell();
	$oCell->createTextRun($DATE[$A])->getFont()->setSize(14);
	$oCell->createTextRun(' / ')->getFont()->setSize(14);
	$oCell->createTextRun($TIME[$A])->getFont()->setSize(14);

	// Add row
	$row = $shape->createRow();
	$row->setHeight(60);
	$row->getFill()->setFillType(Fill::FILL_SOLID)
				   ->setRotation(90)
				   ->setStartColor(new Color('ffffff'))
				   ->setEndColor(new Color('ffffff'));
	$oCell = $row->nextCell();
	$oCell->createTextRun('DETECTION')->getFont()->setBold(true)->setSize(16);
	$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(30);
	$oCell = $row->nextCell();
	$oCell->createTextRun($DETECTION[$A])->getFont()->setSize(14);

	// Add a file drawing (GIF) to the slide
	$shape = new Drawing\File();
	$shape->setName('PHPPresentation logo')
		->setDescription('PHPPresentation logo')
		->setPath("image/image$A.jpg")
		->setHeight(380)
		->setOffsetX(450)
		->setOffsetY(200);
	$currentSlide->addShape($shape);
// Add a file drawing (GIF) to the slide
	$shape = new Drawing\File();
	$shape->setName('PHPPresentation logo')
		->setDescription('PHPPresentation logo')
		->setPath("image/footbar.jpg")
		->setHeight(50.5)
		->setOffsetX(0)
		->setOffsetY(666);
	$currentSlide->addShape($shape);
	// Create a shape (text)
	$shape = $currentSlide->createRichTextShape()
		->setHeight(50)
		->setWidth(400)
		->setOffsetX(0)
		->setOffsetY(680);
	$shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
	$textRun = $shape->createTextRun('รายงานผลการดำเนินงานสิ่งปลูกสร้างรุกล้ำแนวเขตเดินสายส่งไฟฟ้า');
	$textRun->getFont()->setBold(true)
		->setSize(14)
		->setColor(new Color( 'FF0868AC' ));
	foreach ($row->getCells() as $cell) {
		$cell->getBorders()->getTop()->setLineWidth(4)
									 ->setLineStyle(Border::LINE_SINGLE)
									 ->setDashStyle(Border::DASH_DASH);
	}
	$A++;
}
?><center>
<?php

echo write($objPHPPresentation, basename(__FILE__, '.php'), $writers);
if (!CLI) {
	include_once 'Footer.php';
}

?></center>