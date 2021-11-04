<?php
ini_set('error_reporting', E_ALL);
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);
ini_set("memory_limit", "512M");
require($_SERVER["DOCUMENT_ROOT"] . "/bitrix/modules/main/include/prolog_before.php");
use Bitrix\Main\Loader;
use \Bitrix\Catalog\Model\Price;

require_once $_SERVER['DOCUMENT_ROOT'].'/excel/vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Helper\Sample;
require_once($_SERVER['DOCUMENT_ROOT']."/excel/Classes/csv.php");

Loader::includeModule("iblock");
Loader::includeModule("sale");
Loader::includeModule("catalog");
Loader::includeModule("currency");
Loader::includeModule("highloadblock");

global $APPLICATION;

$arHLBlock = \Bitrix\Highloadblock\HighloadBlockTable::getList(array("filter" => array('TABLE_NAME'=>"b_hlbd_tsvetasvecheniya")))->fetch();
$obEntity = Bitrix\Highloadblock\HighloadBlockTable::compileEntity($arHLBlock);
$strEntityDataClass = $obEntity->getDataClass();
$resData = $strEntityDataClass::GetList(array(
    'select' => array('*'),
    'order' => array('ID' => 'ASC')
));
while ($arItem = $resData->Fetch()) {
    $arColors[$arItem["UF_XML_ID"]] = $arItem;
}

$arSections = array();
$arItems = array();
$arID = array();
$arSelect = Array("ID", "NAME", "CODE", "IBLOCK_SECTION_ID", "PREVIEW_PICTURE", "DETAIL_PAGE_URL", "IBLOCK_ID");
$arFilter = Array("IBLOCK_ID"=>18, "ACTIVE" => "Y");
$ob = CIBlockElement::GetList(Array("SORT" => "ASC"), $arFilter, false, array(), $arSelect);
while($res = $ob->GetNextElement()) {
    $fields = $res->GetFields();
    $props = $res->GetProperties();
    $res = $fields;
    foreach ($props as $prop){
        if (is_array($prop["VALUE"])) {
            if (count($prop["VALUE"]) > 1) {
                $res["PROPERTY_".strtoupper($prop["CODE"])."_VALUE"] = $prop["VALUE"];
            } else {
                $res["PROPERTY_".strtoupper($prop["CODE"])."_VALUE"] = $prop["VALUE"][0];
            }
        } else {
            $res["PROPERTY_".strtoupper($prop["CODE"])."_VALUE"] = $prop["VALUE"];
        }
    }
    $res["NAME"] = str_replace('Светодиодный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('Светодиодная ', "", $res["NAME"]);
    $res["NAME"] = str_replace('светодиодный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('светильник ', "", $res["NAME"]);
    $res["NAME"] = str_replace('Светильник ', "", $res["NAME"]);
    $res["NAME"] = str_replace('подвесной ', "", $res["NAME"]);
    $res["NAME"] = str_replace('переносной ', "", $res["NAME"]);
    $res["NAME"] = str_replace('промышленный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('линейный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('Уличный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('низковольтный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('универсальная ', "", $res["NAME"]);
    $res["NAME"] = str_replace('накладной ', "", $res["NAME"]);
    $res["NAME"] = str_replace('компактный ', "", $res["NAME"]);
    $res["NAME"] = str_replace('прожектор ', "", $res["NAME"]);
    $res["NAME"] = str_replace('панель ', "", $res["NAME"]);
    $res["NAME"] = str_replace('для растений ', "", $res["NAME"]);
    $res["NAME"] = str_replace('лампа ', "", $res["NAME"]);
    $res["NAME"] = str_replace('на солнечной батарее ', "", $res["NAME"]);
    $res["NAME"] = str_replace('"Колокол" ', "", $res["NAME"]);
    $res["NAME"] = str_replace('Колокол ', "", $res["NAME"]);
    $res["NAME"] = htmlspecialchars_decode ($res["NAME"]);
    if ($res['PREVIEW_PICTURE']) {
        $res["PIC"] = CFile::ResizeImageGet($res['PREVIEW_PICTURE'], array('width'=>50, 'height'=>50), BX_RESIZE_IMAGE_PROPORTIONAL, true);
        $res["PIC_INFO"] = CFile::GetFileArray($res['PREVIEW_PICTURE']);
        if ($res["PIC_INFO"]["CONTENT_TYPE"] == "image/jpeg") {
            $type = ".jpg";
        } else if ($res["PIC_INFO"]["CONTENT_TYPE"] == "image/png") {
            $type = ".png";
        } else {
            echo "<pre>";
            print_r($res);
            echo "</pre>";
            die();
        }
        if (!copy($_SERVER["DOCUMENT_ROOT"].$res["PIC"]["src"], $_SERVER["DOCUMENT_ROOT"]."/excel/img/".$res["ID"].$type)) {
            echo "<pre>";
            print_r($res["PIC_INFO"]);
            echo "</pre>";
            echo "<pre>";
            print_r("Не сработало копирование файла");
            echo "</pre>";
            die();
        }
        $res["PIC_SRC"] = $_SERVER["DOCUMENT_ROOT"]."/excel/img/".$res["ID"].$type;
    } else {
        $res["PIC_SRC"] = $_SERVER["DOCUMENT_ROOT"]."/bitrix/templates/aspro_next/images/no_photo_medium.png";
    }
    if (!in_array($res["IBLOCK_SECTION_ID"], $arSections)) {
        $arSections[] = $res["IBLOCK_SECTION_ID"];
    }
    if (mb_strtolower($res["PROPERTY_PLASHKA_VALUE"]) == "распродажа") {
        $res["SALE"] = true;
    }
    $arID[] = $res["ID"];
    $arItems[$res["IBLOCK_SECTION_ID"]][$res["ID"]] = $res;
}
$arSelect = Array("ID", "NAME", "IBLOCK_ID", "PRICE_1");
$arFilter = Array("IBLOCK_ID"=>21, "ACTIVE" => "Y", "PROPERTY_CML2_LINK" => $arID);
$ob = CIBlockElement::GetList(Array(), $arFilter, false, array(), $arSelect);
while($res = $ob->GetNextElement()) {
    $fields = $res->GetFields();
    $props = $res->GetProperties();
    $res = $fields;
    foreach ($props as $prop){
        if (is_array($prop["VALUE"])) {
            if (count($prop["VALUE"]) > 1) {
                $res["PROPERTY_".strtoupper($prop["CODE"])."_VALUE"] = $prop["VALUE"];
            } else {
                $res["PROPERTY_".strtoupper($prop["CODE"])."_VALUE"] = $prop["VALUE"][0];
            }
        } else {
            $res["PROPERTY_".strtoupper($prop["CODE"])."_VALUE"] = $prop["VALUE"];
        }
    }
    $res["COLOR"] = $arColors[$res["PROPERTY_COLOR_VALUE"]]["UF_NAME"];
    foreach ($arItems as $key=>$item) {
        if (key_exists($res["PROPERTY_CML2_LINK_VALUE"], $item)) {
            $arItems[$key][$res["PROPERTY_CML2_LINK_VALUE"]]["TP"][] = $res;
        }
    }
}

$arResult[] = $arItems[126];
$arResult[] = $arItems[141];
$arResult[] = $arItems[139];
$arResult[] = $arItems[136];
$arResult[] = $arItems[140];
$arResult[] = $arItems[133];
$arResult[] = $arItems[128];
$arResult[] = $arItems[135];
$arResult[] = $arItems[145];
$arResult[] = $arItems[119];
$arResult[] = $arItems[138];
$arResult[] = $arItems[144];
$arResult[] = $arItems[137];
$arResult[] = $arItems[125];
$arResult[] = $arItems[127];
$arResult[] = $arItems[142];
$arResult[] = $arItems[143];
$arResult[] = $arItems[131];
$arResult[] = $arItems[106];
$arResult[] = $arItems[134];
$arResult[] = $arItems[112];
$arResult[] = $arItems[107];
$arResult[] = $arItems[111];
$arResult[] = $arItems[115];
$arResult[] = $arItems[110];
$arResult[] = $arItems[124];
$arResult[] = $arItems[113];
$arResult[] = $arItems[122];
$arResult[] = $arItems[116];
$arResult[] = $arItems[109];
//$arResult[] = $arItems[114];
$arResult[] = $arItems[108];
$arResult[] = $arItems[123];
$arResult[] = $arItems[121];
$arResult[] = $arItems[130];

$arFilter = array(
    "CURRENCY" => "USD",
    'DATE_RATE' => date('d.m.Y')
);
$by = "date";
$order = "desc";

$rate = CCurrencyRates::GetList($by, $order, $arFilter)->Fetch();
$dollar = ceil($rate["RATE"]*1.05*100)/100;

$massive = array();
foreach ($arResult as $sec=>$arr) {
    foreach ($arr as $item) {
        $arMas = array();
        $i = 0;
        foreach ($item["TP"] as $tp) {
            if ($i == 0) {
                $price = intval($tp["PRICE_1"]);
                $price_d1 = false;
                $price_d2 = false;
                $price_d3 = false;
                $price_e1 = false;
                $price_e2 = false;
                $type_price = $item["PROPERTY_PRICE_TYPE_VALUE"];
                if ($type_price == 1 || !$type_price) {
//                    $price_d1 = $price*(100-30)/100;
//                    $price_d2 = $price*(100-40)/100;
//                    $price_d3 = $price*(100-50)/100;
//                    $price_e1 = $price*(100-40)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*5*$dollar;
//                    $price_e2 = $price*(100-40)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*6.5*$dollar;
                    $price_d1 = $price*(100-3)/100;
                    $price_d2 = $price*(100-4)/100;
                    $price_d3 = $price*(100-5)/100;
                    $price_e1 = $price/1.22*(100-5)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*5.6*$dollar;
                    $price_e2 = $price/1.22*(100-5)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*6.5*$dollar;
                } else if ($type_price == 2) {
                    $price_d1 = $price*(100-6)/100;
                    $price_d2 = $price*(100-8)/100;
                    $price_d3 = $price*(100-10)/100;
                    $price_e1 = $price/1.22*(100-10)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*5.6*$dollar;
                    $price_e2 = $price/1.22*(100-10)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*6.5*$dollar;
                } else if ($type_price == 3) {
                    $price_d1 = $price*(100-9)/100;
                    $price_d2 = $price*(100-12)/100;
                    $price_d3 = $price*(100-15)/100;
                    $price_e1 = $price/1.22*(100-15)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*5.6*$dollar;
                    $price_e2 = $price/1.22*(100-15)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*6.5*$dollar;
                } else if ($type_price == 4) {
                    $price_d1 = $price*(100-13)/100;
                    $price_d2 = $price*(100-17)/100;
                    $price_d3 = $price*(100-20)/100;
                    $price_e1 = $price/1.22*(100-20)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*5.6*$dollar;
                    $price_e2 = $price/1.22*(100-20)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*6.5*$dollar;
                } else if ($type_price == 5) {
                    $price_d1 = $price*(100-17)/100;
                    $price_d2 = $price*(100-22)/100;
                    $price_d3 = $price*(100-25)/100;
                    $price_e1 = $price/1.22*(100-25)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*5.6*$dollar;
                    $price_e2 = $price/1.22*(100-25)/100 + $item["PROPERTY_WEIGHT_VALUE"]/1000*6.5*$dollar;
                }
                if (intval($price_e1) < intval($price_d2)) {
                    $price_e1 = $price_d2;
                }
                if (intval($price_e2) < intval($price_d2)) {
                    $price_e2 = $price_d2;
                }
            }
            $arMas[0] .= $tp["PROPERTY_ARTICLE_VALUE"].":";
            if (substr_count($tp["COLOR"], "Холодный белый")) {
                $arMas[1] .= "ffffff:";
            } else if (substr_count($tp["COLOR"], "Нейтральный белый")) {
                $arMas[1] .= "ffff00:";
            } else if (substr_count($tp["COLOR"], "Теплый белый")) {
                $arMas[1] .= "ffa500:";
            } else {
                $arMas[1] .= "ffffff:";
            }
            $i++;
        }
        $arMas[0] = substr($arMas[0], 0, -1);
        $arMas[1] = substr($arMas[1], 0, -1);
        $arMas[2] = $item["PIC_SRC"];
        $arMas[3] = $price;
        $arMas[4] = $item["NAME"]."|https://favouritestyle.ru".$item["DETAIL_PAGE_URL"];
        $arMas[5] = $item["PROPERTY_POWER_VALUE"];
        $arMas[6] = $item["PROPERTY_LUMINOUS_FLUX_VALUE"];
        $arMas[7] = $item["PROPERTY_GABARIT_VALUE"];
        $arMas[8] = $item["PROPERTY_WEIGHT_VALUE"];
        $arMas[9] = $item["PROPERTY_STEPEN_ZAW_VALUE"];
        $arMas[10] = htmlspecialchars_decode($item["PROPERTY_CRI_VALUE"]);
        $arMas[11] = $item["PROPERTY_UGOL_SVET_PU4_VALUE"];
        $arMas[12] = htmlspecialchars_decode($item["PROPERTY_KOEF_PULS_SVET_POTOK_VALUE"]);
        $arMas[13] = htmlspecialchars_decode($item["PROPERTY_COFFICIENT_POWER_VALUE"]);
        $arMas[14] = $item["PROPERTY_GARANTIYA_VALUE"];
        $arMas[15] = $item["PROPERTY_SROK_EKSPL_VALUE"];
        $arMas[16] = $item["SALE"];
        $arMas[] = ceil($price_d1*100)/100;
        $arMas[] = ceil($price_d2*100)/100;
        $arMas[] = ceil($price_d3*100)/100;
        $arMas[] = ceil($price_e1*100)/100;
        $arMas[] = ceil($price_e2*100)/100;
        $massive[$sec][] = $arMas;
    }
}

$helper = new Sample();
$path_parts = pathinfo($_SERVER['SCRIPT_FILENAME']); // определяем директорию скрипта
chdir($path_parts['dirname']);

$reader = IOFactory::createReader('Xlsx');
$spreadsheet = $reader->load($_SERVER['DOCUMENT_ROOT'].'/excel/ftp/shablon.xlsx');


$i=0;
foreach ($massive as $vkladka) {

    $start = 4;

    $sprtec = $spreadsheet->setActiveSheetIndex($i);
    $sprtec->getStyle('C' . ($start-1))->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setRGB("538DD5");
    $i++;
    $sprtec->getStyle('C7:R7')->getBorders()
        ->getAllBorders()
        ->setBorderStyle(Border::BORDER_THIN)
        ->setColor(new Color('FFFFFFFF'));
    foreach ($vkladka as $str=>$t) {

        $s = $start;
        $c=0;

        $color=explode(':',$t[1]);
        $arts=explode(':',$t[0]);
        foreach ($color as $j=>$col) {
            $sprtec->setCellValue('A' . $start, $arts[$j]);
            $sprtec->getStyle('A' . $start)
                ->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB($col);
            $c++;
            $sprtec->getStyle('A' . $start)->getBorders()
                ->getAllBorders()
                ->setBorderStyle(Border::BORDER_THIN)
                ->setColor(new Color('000000'));
            $sprtec->getStyle('B' . $start)->getBorders()
                ->getVertical()
                ->setBorderStyle(Border::BORDER_THIN)
                ->setColor(new Color('000000'));
            $start++;
        }

        if (count($color) == 2) {
            $sprtec->mergeCells('A' . ($s+1) . ':A' . ($start-1));
            $sprtec->getRowDimension($s+1)->setRowHeight(22);
            $sprtec->getRowDimension($s)->setRowHeight(23);
        }
        if (count($color) == 1) {
            $sprtec->mergeCells('A' . $s . ':A' . ($start - 1));
            $sprtec->getRowDimension($s)->setRowHeight(45);
        }

        $sprtec->getStyle('A'.($s+2))->getBorders()
            ->getVertical()
            ->setBorderStyle(Border::BORDER_THIN)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('A'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));

        $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
        $drawing->setName('Logo');
        $drawing->setDescription('Logo');
        $drawing->setCoordinates('B' . $s);
        $drawing->setPath($t[2]);
        $drawing->setWidth(50);
        $drawing->setOffsetX(20);
        $drawing->setOffsetY(5);
        $drawing->setWorksheet($spreadsheet->getActiveSheet());
        $sprtec->getStyle('B'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));

        $sprtec->mergeCells('C' . $s . ':C' . ($start-1))->setCellValue('C' . $s, $t[3]);
        $sprtec->getStyle('C'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('C' . $s)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB("538DD5");

        $sprtec->mergeCells('D' . $s . ':D' . ($start-1))->setCellValue('D' . $s, $t[17]);
        $sprtec->getStyle('D'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('D' . $s)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB("CCC0DA");

        $sprtec->mergeCells('E' . $s . ':E' . ($start-1))->setCellValue('E' . $s, $t[18]);
        $sprtec->getStyle('E'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('E' . $s)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB("E6B8B7");

        $sprtec->mergeCells('F' . $s . ':F' . ($start-1))->setCellValue('F' . $s, $t[19]);
        $sprtec->getStyle('F'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('F' . $s)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB("FFC000");

        $sprtec->mergeCells('G' . $s . ':G' . ($start-1))->setCellValue('G' . $s, $t[20]);
        $sprtec->getStyle('G'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('G' . $s)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB("FFFF99");

        $sprtec->mergeCells('H' . $s . ':H' . ($start-1))->setCellValue('H' . $s, $t[21]);
        $sprtec->getStyle('H'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->getStyle('H' . $s)->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setRGB("FFFF00");

        $name=explode('|',$t[4]);
//       echo '=HYPERLINK("'.$name[1].'";"'.$name[1].'")';
        $sprtec->setCellValue('I' . $s, $name[0]);
        $sprtec->getCell('I' . $s)->getHyperlink()->setUrl($name[1]);
        $sprtec->mergeCells('I' . $s . ':I' . ($start-1));
        $sprtec->getStyle('I' . $s)->getFont()->getColor()->setRGB('0000ff');
        if ($t[16]) {
            $sprtec->getStyle('I' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("ff0000");
        } else if ($str % 2 == 1) {
            $sprtec->getStyle('I' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->getStyle('I' . $s)->getFont()->setUnderline(true);
        $sprtec->getStyle('I'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        $sprtec->mergeCells('J' . $s . ':J' . ($start-1))->setCellValue('J' . $s, $t[5]);
        $sprtec->getStyle('J'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('J' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('K' . $s . ':K' . ($start-1))->setCellValue('K' . $s, $t[6]);
        $sprtec->getStyle('K'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('K' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('L' . $s . ':L' . ($start-1))->setCellValue('L' . $s, $t[7]);
        $sprtec->getStyle('L'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('L' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('M' . $s . ':M' . ($start-1))->setCellValue('M' . $s, $t[8]);
        $sprtec->getStyle('M'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('M' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('N' . $s . ':N' . ($start-1))->setCellValue('N' . $s, $t[9]);
        $sprtec->getStyle('N'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('N' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('O' . $s . ':O' . ($start-1))->setCellValue('O' . $s, $t[10]);
        $sprtec->getStyle('O'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('O' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('P' . $s . ':P' . ($start-1))->setCellValue('P' . $s, $t[11]);
        $sprtec->getStyle('P'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('P' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('Q' . $s . ':Q' . ($start-1))->setCellValue('Q' . $s, $t[12]);
        $sprtec->getStyle('Q'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('Q' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('R' . $s . ':R' . ($start-1))->setCellValue('R' . $s, $t[13]);
        $sprtec->getStyle('R'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('R' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('S' . $s . ':S' . ($start-1))->setCellValue('S' . $s, $t[14]);
        $sprtec->getStyle('S'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('S' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }
        $sprtec->mergeCells('T' . $s . ':T' . ($start-1))->setCellValue('T' . $s, $t[15]);
        $sprtec->getStyle('T'.($start-1))->getBorders()
            ->getBottom()
            ->setBorderStyle(Border::BORDER_MEDIUM)
            ->setColor(new Color('000000'));
        if ($str % 2 == 1) {
            $sprtec->getStyle('T' . $s)->getFill()
                ->setFillType(Fill::FILL_SOLID)
                ->getStartColor()->setRGB("D9D9D9");
        }

    }
   $sprtec->getStyle('B4:U' . ($start-1))->getBorders()
        ->getVertical()
        ->setBorderStyle(Border::BORDER_THIN)
        ->setColor(new Color('000000'));
    $sprtec->getStyle('A4:U' . ($start-1))->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
    $sprtec->getStyle('A4:U' . ($start-1))->getAlignment()->setWrapText(true);

}

$sheet = $spreadsheet->setActiveSheetIndex($i);
$sheet->mergeCells("A1:AP55");
$sheet->setCellValue("A1", "");

$i++;
$sheet = $spreadsheet->setActiveSheetIndex($i);
$sheet->mergeCells("A1:AP55");
$sheet->setCellValue("A1", "");

$main_page = $i++;
$sheet = $spreadsheet->setActiveSheetIndex($i);
$sheet->mergeCells("A1:AP55");
$sheet->setCellValue("A1", "");

$i++;
$sheet = $spreadsheet->setActiveSheetIndex($i);
$sheet->mergeCells("A1:AP55");
$sheet->setCellValue("A1", "");

$i++;
$sheet = $spreadsheet->setActiveSheetIndex($i);
$sheet->mergeCells("A1:AP55");
$sheet->setCellValue("A1", "");

$i++;
$sheet = $spreadsheet->setActiveSheetIndex($i);
$sheet->mergeCells("A1:AP55");
$sheet->setCellValue("A1", "");

$spreadsheet->setActiveSheetIndex($main_page);

// Save
$oWriter = IOFactory::createWriter($spreadsheet, 'Xlsx');
$oWriter->save($_SERVER['DOCUMENT_ROOT']."/excel/files/price-ledfavourite-".date("d-m-Y").".xlsx");

$file_url = 'https://favouritestyle.ru/excel/files/price-ledfavourite-'.date("d-m-Y").'.xlsx';

?>
<!doctype html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport"
          content="width=device-width, user-scalable=no, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Excel</title>
</head>
<body>
<p>
    Файл сформирован
</p>
<p>
    <a href="<?=$file_url?>">Скачать</a>
</p>
</body>
</html>