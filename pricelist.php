<?php
include_once($_SERVER['DOCUMENT_ROOT']."/xlsxwriter/xlsxwriter.class.php");
require_once($_SERVER['DOCUMENT_ROOT']."/xlsxwriter/xlsxwriterplus.class.php");

global $nn;
$nn = 1;
global $color_white; // #ffffff
global $color_light; // #fafafa
global $color_dark;  // #303030
global $color_blue;  // #22b8f0
global $color_black; // #000000
$color_white = '#ffffff';
$color_light = '#fafafa';
$color_dark = '#303030';
$color_blue = '#22b8f0';
$color_black = '#000000';

global $fb_tmpdir;
global $euro_rate;
global $today_date;
$fb_tmpdir = "/tmp";
$today_date = date("d.m.Y H:i" );

$filename = "ADLPricelist.xlsx";

$preheader = array(
        array( ' ', ' ', ' ', ' ' ),
        array( 'ООО "ПК АДЛ Электроникс"', ' ', ' ', ' ' ),
        array( 'Санкт-Петербург, ул.Ольги Берггольц, д.35А', ' ', ' ', ' ' ),
        array( 'Россия', ' ', ' ', ' ' ),
        array( '(812) 568-18-91', ' ', ' ', ' ' ),
        array( ' ', ' ', ' ', ' ' ),
);
$preheader1 = array(
        array( '' ),
        array( '', 'Прайслист от '.date("d.m.Y" ), '1 EUR=', $euro_rate ),
        array( '' ),
);

$headertypes = array(
        '№'=>'string',
        'Описание'=>'string',
        'Цена, EUR'=>'euro',
        'Цена, RUB*'=>'string',
);

$preheader1types = array(
        '№'=>'string',
        'Описание'=>'string',
        'Цена, EUR'=>'string',
        'Цена, RUB'=>'# ##0.0000 "₽";-# ##0.0000 "₽"',
);

$space = array( '' );
$footnote_xls = array(
        array( "*) Цена в рублях = цена в евро по курсу ЦБ на текущий день +2%" ),
        array( "Курс евро ЦБ России на ".$today_date.": "."1"." EUR = ".$euro_rate." RUB" )
);

$disclaimer = array( "Прайслист содержит информацию о продукции, сервисном обслуживании, рекламных программах и мероприятиях компании \"АДЛ Электроникс\".  ".
        "Все содержащиеся в прайслисте сведения носят исключительно информационный характер и не является исчерпывающими.  ".
        "Указанные цены являются рекомендованными розничными ценами и могут отличаться от действительных цен уполномоченных дилеров.  ".
        "Представленная на сайте информация, касающаяся комплектаций, технических характеристик, внешнего вида, стоимости, условий приобретения продукции,  ".
        "сервисного обслуживания и т. п. может отличаться от действительных характеристик и условий приобретения продукции может быть изменена в любое время ".
        "без предварительного уведомления. Представленные в каталоге изображения изделий, являются графическим представлением изделия и могут ".
        "незначительно отличатся от оригинала. Представленная на сайте информация о продукции не означает, что последняя есть в наличии для продажи. ".
        "Более подробную и точную информацию можно получить в офисе компании \"АДЛ Электроникс\" или у официальных дилеров.");
$copy =    array( '© 2020 АДЛ Электроникс. Все права защищены.');

global $writer;
global $sheet;

$writer = new XLSWriterPlus();
$writer->setAuthor('ADL Electronics Ltd.');
$sheet = 'Прайслист';

$writer->setColumnsOptions( $sheet, $headertypes, $col_options=[ 'widths'=>[6,60,12,14] ]);
foreach($preheader as $row) {
   $writer->writeSheetRow( $sheet, $row, $col_options=[ 'color'=>$color_blue, 'fill'=>$color_dark ]);
     $current_row = $writer->countSheetRows( $sheet );
     $writer->markMergedCell($sheet, $current_row-1, 0, $current_row-1, 3 );
}

$writer->setColumnsOptions( $sheet, $preheader1types,  $col_options=[ 'widths'=>[6,60,12,14] ]);
foreach($preheader1 as $row) {
        $writer->writeSheetRow( $sheet, $row,
                $styles=[ ['halign'=>'left'],
                                    ['halign'=>'center','font-size'=>12, 'font-style'=>'bold' ],
                                    ['halign'=>'right'],
                                    ['halign'=>'right'] ]);
}
$writer->addImage( realpath($_SERVER['DOCUMENT_ROOT'].'/images/Blue_mini_logo_xlsx.png'), 1, ['startColNum' => 2, 'startRowNum' => 1, 'endColNum' => 3, 'endRowNum' => 4 ] );
$writer->writeSheetHeader( $sheet, $headertypes,
        $styles=[ ['halign'=>'center','border'=>'left,right,top,bottom'],
                            ['halign'=>'center','border'=>'left,right,top,bottom'],
                            ['halign'=>'center','border'=>'left,right,top,bottom'],
                            ['halign'=>'center','border'=>'left,right,top,bottom'] ], 1 );
            // $col_options=[ 'height'=>24, 'halign'=>'center', 'valign'=>'center', 'color'=>'', 'fill'=>'', 'border'=>'left,right,top,bottom' ], 1 );


function getCategory( $parent_id, $order, $link, $depth )
{
        global $euro_rate;
        global $nn;
        global $uid;
        global $writer;
        global $sheet;
        global $color_white; // #ffffff
        global $color_light; // #fafafa
        global $color_dark;  // #303030
        global $color_blue;  // #22b8f0
        global $color_black; // #000000

        $cat_query = "SELECT name, id, ordering, alias FROM #__k2_categories WHERE published=1 AND trash=0 AND parent=".$parent_id." ORDER BY ordering;";
        $db->setQuery($cat_query);
        $rows = $db->loadObjectList();
        foreach( $rows as $row )
        {
                if( $depth == 0 )
                {
                        $cat_num = "";
                        $suffix = "";
                        $row_font = $color_white;
                        $row_back = $color_dark;
                }
                else
                {
                        $cat_num = substr( $order, 2 );
                        $suffix = $row->ordering;
                        switch( $depth )
                        {
                                case 1:
                                        $row_font = $color_blue;
                                        $row_back = $color_dark;
                                        break;
                                case 2:
                                        $row_font = $color_black;
                                        $row_back = $color_blue;
                                        break;
                                case 3:
                                        $row_font = $color_black;
                                        $row_back = $color_light;
                                        break;
                        }
                }
                // XLSX Category
                if( $cat_num != '' || $suffix != '' )
                {
                        $xrow = array( "".$cat_num.$suffix."", $row->name, ' ', ' ' );
                        $writer->writeSheetRow( $sheet, $xrow, $row_options = ['height'=>24, 'valign'=>'center', 'color'=>$row_font, 'fill'=>$row_back, 'border'=>'left,right,top,bottom']);
                        $current_row = $writer->countSheetRows( $sheet );
                        $writer->markMergedCell($sheet, $current_row-1, 1, $current_row-1, 3 );
                }
                // XLSX

                $item_query = "SELECT title, price, alias, id";
                $item_query .= " FROM #__k2_items WHERE published=1 AND trash=0 AND catid=".$row->id.";";
                $db->setQuery($item_query);
                $items = $db->loadObjectList();
                $even = "even";

                foreach( $items as $item )
                {
                        $ue = $item->price;
                        $rub = $item->price * $euro_rate *1.02 ;

                        if( $ue == 0 )
                        {
                                $ue = "";
                                $rub = "по запросу";
                        }
                        else
                        {
                                $ue= number_format($ue,2,'.',' ');
                                $rub = number_format($rub,2,'.',' ');
                        }

                        // XLSX Item

                        $current_row = $writer->countSheetRows( $sheet )+1;

                        if( $rub == 'по запросу' ) {
                                $col_types = array(
                                        '№'=>'string',
                                        'Описание'=>'string',
                                        'Цена, EUR'=>'string',
                                        'Цена, RUB'=>'string',
                                );

                                $xrow = array( $nn, $item->title, $rub, '');
                                $writer->setColumnsOptions( $sheet, $col_types, $col_options=[ 'widths'=>[6,60,12,14] ]);
                                $writer->writeSheetRow( $sheet, $xrow,
                                        $styles=[ ['halign'=>'','border'=>'left,right,top,bottom'],
                                                            ['halign'=>'','border'=>'left,right,top,bottom'],
                                                            ['halign'=>'center','border'=>'left,right,top,bottom'],
                                                            ['halign'=>'center','border'=>'left,right,top,bottom'] ] );
                                $writer->markMergedCell($sheet, $current_row-1, 2, $current_row-1, 3 );
                        }
                        else {
                                $rub = str_replace( ' ', '', $rub );
                                $col_types = array(
                                        '№'=>'string',
                                        'Описание'=>'string',
                                        'Цена, EUR'=>'euro',
                                        'Цена, RUB'=>'# ##0.00 "₽";-# ##0.00 "₽"',
                                );
                                $xrow = array( $nn, $item->title, str_replace( ' ', '', $ue), '=C'.$current_row.'*1.02*$D$8');
                                $writer->writeSheetRow( $sheet, $xrow, $row_options = [/*'height'=>24,*/ 'valign'=>'center', 'fill'=>$bgcolor, 'border'=>'left,right,top,bottom'], $col_types );
                        }
                        // XLSX

                        $nn += 1;
                        if( $even == "even" )
                                $even = "odd";
                        else
                                $even = "even";
                }
                getCategory( $row->id, $order.$row->ordering.".", $link."/".$row->alias, $depth+1 );
        }
}

getCategory( 0, "", "index.php", 0 );

$writer->writeSheetRow( $sheet, $space);
foreach($footnote_xls as $row) {
   $writer->writeSheetRow( $sheet, $row);
     $current_row = $writer->countSheetRows( $sheet );
     $writer->markMergedCell($sheet, $current_row-1, 0, $current_row-1, 3 );
}
$writer->writeSheetRow( $sheet, $space);

$writer->writeSheetRow( $sheet, $disclaimer, $row_options = ['height'=>146, 'valign'=>'distributed']);
$current_row = $writer->countSheetRows( $sheet );
$writer->markMergedCell($sheet, $current_row-1, 0, $current_row-1, 3 );

$writer->writeSheetRow( $sheet, $space);
$writer->writeSheetRow( $sheet, $copy);

$writer->writeToFile( $_SERVER['DOCUMENT_ROOT']."/tmp/".$filename );
?>

